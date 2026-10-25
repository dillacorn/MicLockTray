using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace MicLockTray;

// All endpoint discovery is on demand or triggered by Core Audio notifications; no polling.
internal static class DeviceCatalog
{
    internal sealed record Microphone(string Id, string Name, int CurrentPercent, bool IsDefault);

    private static readonly Guid EndpointVolumeId = new("5CDF2C82-841E-4546-9722-0CF74078229A");
    private static readonly Guid FriendlyNamePropertyId = new("A45C254E-DF1C-4EFD-8020-67D146A850E0");

    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPropertyStore
    {
        [PreserveSig] int GetCount(out uint count);
        [PreserveSig] int GetAt(uint index, out CoreAudio.PropertyKey key);
        [PreserveSig] int GetValue(ref CoreAudio.PropertyKey key, out PropVariant value);
        [PreserveSig] int SetValue(ref CoreAudio.PropertyKey key, ref PropVariant value);
        [PreserveSig] int Commit();
    }

    [StructLayout(LayoutKind.Explicit, Size = 24)]
    private struct PropVariant
    {
        [FieldOffset(0)] public ushort Type;
        [FieldOffset(8)] public IntPtr Pointer;
    }

    [DllImport("ole32.dll")]
    private static extern int PropVariantClear(ref PropVariant value);

    public static IReadOnlyList<Microphone> GetActiveMicrophones()
    {
        var result = new List<Microphone>();
        try
        {
            var enumerator = (CoreAudio.IMMDeviceEnumerator)new CoreAudio.MMDeviceEnumerator();
            var defaults = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var role in new[] { CoreAudio.ERole.eConsole, CoreAudio.ERole.eMultimedia, CoreAudio.ERole.eCommunications })
            {
                if (enumerator.GetDefaultAudioEndpoint(CoreAudio.EDataFlow.eCapture, role, out var device) == 0 &&
                    device != null && TryGetId(device, out string id))
                    defaults.Add(id);
            }

            // DEVICE_STATE_ACTIVE. Disconnected devices are retained by Settings, not Windows' live list.
            if (enumerator.EnumAudioEndpoints(CoreAudio.EDataFlow.eCapture, 0x1, out var collection) != 0 ||
                collection is null || collection.GetCount(out uint count) != 0)
                return result;

            for (uint i = 0; i < count; i++)
            {
                if (collection.Item(i, out var device) != 0 || device is null ||
                    !TryGetId(device, out string id))
                    continue;

                string name = ReadFriendlyName(device) ?? id;
                int volume = 100;
                try
                {
                    Guid iid = EndpointVolumeId;
                    if (device.Activate(ref iid, 0x1, IntPtr.Zero, out object instance) == 0 &&
                        instance is CoreAudio.IAudioEndpointVolume endpoint &&
                        endpoint.GetMasterVolumeLevelScalar(out float scalar) == 0)
                        volume = Math.Clamp((int)Math.Round(scalar * 100f), 1, 100);
                }
                catch (Exception ex) { Debug.WriteLine($"MicLockTray volume read failed for {id}: {ex}"); }

                result.Add(new Microphone(id, name, volume, defaults.Contains(id)));
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray microphone enumeration failed: {ex}");
        }

        return result;
    }

    private static bool TryGetId(CoreAudio.IMMDevice device, out string id)
    {
        id = string.Empty;
        IntPtr pointer = IntPtr.Zero;
        try
        {
            if (device.GetId(out pointer) != 0 || pointer == IntPtr.Zero) return false;
            id = Marshal.PtrToStringUni(pointer) ?? string.Empty;
            return id.Length != 0;
        }
        finally
        {
            if (pointer != IntPtr.Zero) Marshal.FreeCoTaskMem(pointer);
        }
    }

    private static string? ReadFriendlyName(CoreAudio.IMMDevice device)
    {
        IntPtr pointer = IntPtr.Zero;
        try
        {
            // STGM_READ = 0. IMMDevice.OpenPropertyStore returns a COM IPropertyStore pointer.
            if (device.OpenPropertyStore(0, out pointer) != 0 || pointer == IntPtr.Zero)
                return null;

            var store = (IPropertyStore)Marshal.GetObjectForIUnknown(pointer);
            var key = new CoreAudio.PropertyKey { formatId = FriendlyNamePropertyId, propertyId = 14 };
            if (store.GetValue(ref key, out var value) != 0) return null;
            try
            {
                return value.Type switch
                {
                    31 => Marshal.PtrToStringUni(value.Pointer), // VT_LPWSTR
                    8 => Marshal.PtrToStringBSTR(value.Pointer), // VT_BSTR
                    _ => null
                };
            }
            finally { _ = PropVariantClear(ref value); }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"MicLockTray device name read failed: {ex}");
            return null;
        }
        finally
        {
            if (pointer != IntPtr.Zero) Marshal.Release(pointer);
        }
    }
}
