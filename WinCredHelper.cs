using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

/// <summary>
/// Thin wrapper around the Windows Credential Manager (advapi32 CredWrite/CredRead).
/// Stores and retrieves opaque secrets under a generic credential target name.
/// </summary>
public static class WinCredHelperV2
{
    private const uint CRED_TYPE_GENERIC         = 1;
    private const uint CRED_PERSIST_LOCAL_MACHINE = 2;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct CREDENTIAL
    {
        public uint   Flags;
        public uint   Type;
        [MarshalAs(UnmanagedType.LPWStr)] public string TargetName;
        [MarshalAs(UnmanagedType.LPWStr)] public string Comment;
        public long   LastWritten;
        public uint   CredentialBlobSize;
        public IntPtr CredentialBlob;
        public uint   Persist;
        public uint   AttributeCount;
        public IntPtr Attributes;
        [MarshalAs(UnmanagedType.LPWStr)] public string TargetAlias;
        [MarshalAs(UnmanagedType.LPWStr)] public string UserName;
    }

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CredWriteW(ref CREDENTIAL credential, uint flags);

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CredReadW(string target, uint type, uint flags, out IntPtr credential);

    [DllImport("advapi32.dll", SetLastError = true, EntryPoint = "CredFree")]
    private static extern void CredFree(IntPtr buffer);

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CredDeleteW(string target, uint type, uint flags);

    /// <summary>Saves (or overwrites) a secret string under the given target name.</summary>
    public static void SaveSecret(string target, string secret)
    {
        byte[] blob = Encoding.Unicode.GetBytes(secret);
        IntPtr ptr  = Marshal.AllocHGlobal(blob.Length);
        try
        {
            Marshal.Copy(blob, 0, ptr, blob.Length);
            CREDENTIAL cred = new CREDENTIAL
            {
                Type               = CRED_TYPE_GENERIC,
                TargetName         = target,
                UserName           = "AltTextGenerator",
                CredentialBlobSize = (uint)blob.Length,
                CredentialBlob     = ptr,
                Persist            = CRED_PERSIST_LOCAL_MACHINE
            };
            if (!CredWriteW(ref cred, 0))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        finally
        {
            Marshal.FreeHGlobal(ptr);
        }
    }

    /// <summary>Returns the stored secret, or null if the target does not exist.</summary>
    public static string GetSecret(string target)
    {
        IntPtr ptr;
        if (!CredReadW(target, CRED_TYPE_GENERIC, 0, out ptr))
            return null;
        try
        {
            CREDENTIAL cred = (CREDENTIAL)Marshal.PtrToStructure(ptr, typeof(CREDENTIAL));
            if (cred.CredentialBlobSize == 0) return string.Empty;
            byte[] blob = new byte[(int)cred.CredentialBlobSize];
            Marshal.Copy(cred.CredentialBlob, blob, 0, blob.Length);
            return Encoding.Unicode.GetString(blob);
        }
        finally
        {
            CredFree(ptr);
        }
    }

    /// <summary>Returns true when a credential with the given target name exists.</summary>
    public static bool Exists(string target)
    {
        IntPtr ptr;
        if (!CredReadW(target, CRED_TYPE_GENERIC, 0, out ptr))
            return false;
        CredFree(ptr);
        return true;
    }

    /// <summary>Deletes a credential. Silently succeeds if the target does not exist.</summary>
    public static void Delete(string target)
    {
        CredDeleteW(target, CRED_TYPE_GENERIC, 0);
    }
}
