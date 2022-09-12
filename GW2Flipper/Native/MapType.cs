namespace GW2Flipper.Native;

internal enum MapType : uint
{
    /// <summary>Maps a virtual key to a key code.</summary>
    MAPVK_VK_TO_VSC = 0x00,

    /// <summary>Maps a key code to a virtual key.</summary>
    MAPVK_VSC_TO_VK = 0x01,

    /// <summary>Maps a virtual key to a character.</summary>
    MAPVK_VK_TO_CHAR = 0x02,

    /// <summary>Maps a key code to a virtual key with specified keyboard.</summary>
    MAPVK_VSC_TO_VK_EX = 0x03,

    /// <summary>Maps a virtual key to a key code with specified keyboard.</summary>
    MAPVK_VK_TO_VSC_EX = 0x04,
}
