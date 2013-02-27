using System;
using System.Diagnostics;
using System.IO;

namespace MusicBeeSnarlPlugin
{
    public sealed class VariablesHelper
    {
        private static String playerPath = null;

        private MusicBeePlugin.Plugin.MusicBeeApiInterface mbApiInterface;


        public VariablesHelper(MusicBeePlugin.Plugin.MusicBeeApiInterface mbApiInterface)
        {
            this.mbApiInterface = mbApiInterface;
        }

        public string ReplacePathVariables(String str)
        {
            str = str.ToLowerInvariant();

            if (str.Contains("%musicbee_path%"))
                str = str.Replace("%musicbee_path%", GetPlayerPath());

            return str;
        }

        public String ReplaceNowPlayingVariables(String str)
        {
            str = str.ToLowerInvariant();

            if (str.Contains("%artist%"))
                str = str.Replace("%artist%", mbApiInterface.NowPlaying_GetFileTag(MusicBeePlugin.Plugin.MetaDataType.Artist));
            if (str.Contains("%track_title%"))
                str = str.Replace("%track_title%", mbApiInterface.NowPlaying_GetFileTag(MusicBeePlugin.Plugin.MetaDataType.TrackTitle));
            if (str.Contains("%album%"))
                str = str.Replace("%album%", mbApiInterface.NowPlaying_GetFileTag(MusicBeePlugin.Plugin.MetaDataType.Album));
            if (str.Contains("%track_no%"))
                str = str.Replace("%track_no%", mbApiInterface.NowPlaying_GetFileTag(MusicBeePlugin.Plugin.MetaDataType.TrackNo));
            if (str.Contains("%track_year%"))
                str = str.Replace("%track_year%", mbApiInterface.NowPlaying_GetFileTag(MusicBeePlugin.Plugin.MetaDataType.Year));

            if (str.Contains("%newline%"))
                str = str.Replace("%newline%", "\n");

            return str;
        }

        public static string GetPlayerPath()
        {
            if (playerPath == null)
            {
                using (Process process = Process.GetCurrentProcess())
                {
                    playerPath = Path.GetDirectoryName(process.Modules[0].FileName);
                }
            }
            return playerPath;
        }
    }
}
