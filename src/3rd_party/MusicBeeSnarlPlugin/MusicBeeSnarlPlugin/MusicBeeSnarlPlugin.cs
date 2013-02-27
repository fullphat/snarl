using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using MusicBeeSnarlPlugin;

namespace MusicBeePlugin
{
    public partial class Plugin
    {
        private PluginInfo about = new PluginInfo();
        private MusicBeeApiInterface mbApiInterface;
        private SnarlMusicBeePlayer snarl;
        private VariablesHelper varHelper;

        private Settings settings;
        private String dataPath;
        private String nowPlayingFileUrl = "";
        private String nowPlayingHeadline = "";
        private String nowPlayingText = "";
        private String defaultIconPath = "";


        public PluginInfo Initialise(IntPtr apiInterfacePtr)
        {
            mbApiInterface = (MusicBeeApiInterface)Marshal.PtrToStructure(apiInterfacePtr, typeof(MusicBeeApiInterface));
            about.PluginInfoVersion = PluginInfoVersion;
            about.Name = "Snarl";
            about.Description = "Display Now Playing information with Snarl.";
            about.Author = "Toke Noer";
            about.TargetApplication = ""; //  "Yahoo Messenger";   // current only applies to artwork, lyrics or instant messenger name that appears in the provider drop down selector or target Instant Messenger
            about.Type = PluginType.General;
            about.VersionMajor = 1;
            about.VersionMinor = 0;
            about.Revision = 1;
            about.MinInterfaceVersion = MinInterfaceVersion;
            about.MinApiRevision = MinApiRevision;
            about.ReceiveNotifications = ReceiveNotificationFlags.PlayerEvents;
            about.ConfigurationPanelHeight = 0;   // not implemented yet: height in pixels that musicbee should reserve in a panel for config settings. When set, a handle to an empty panel will be passed to the Configure function

            // save any persistent settings in a sub-folder of this path
            dataPath = mbApiInterface.Setting_GetPersistentStoragePath();

            varHelper = new VariablesHelper(mbApiInterface);

            settings = Settings.Load(dataPath);
            defaultIconPath = varHelper.ReplacePathVariables(settings.DefaultIconPath);

            snarl = new SnarlMusicBeePlayer(mbApiInterface);
            snarl.Register(defaultIconPath);

            return about;
        }

        public bool Configure(IntPtr panelHandle)
        {
            SettingsForm sf = new SettingsForm(settings);
            sf.ShowDialog();

            return true;
        }
       
        // called by MusicBee when the user clicks Apply or Save in the MusicBee Preferences screen.
        // its up to you to figure out whether anything has changed and needs updating
        public void SaveSettings()
        {
            settings.Save(dataPath);
            defaultIconPath = varHelper.ReplacePathVariables(settings.DefaultIconPath);
        }

        // MusicBee is closing the plugin (plugin is being disabled by user or MusicBee is shutting down)
        public void Close(PluginCloseReason reason)
        {
            snarl.Unregister();
        }

        // uninstall this plugin - clean up any persisted files
        public void Uninstall()
        {
            settings.DeleteSettings(dataPath);
        }

        // receive event notifications from MusicBee
        // you need to set about.ReceiveNotificationFlags = PlayerEvents to receive all notifications, and not just the startup event
        public void ReceiveNotification(string sourceFileUrl, NotificationType type)
        {
            // perform some action depending on the notification type
            switch (type)
            {
                case NotificationType.PluginStartup:
                    // perform startup initialization
                    if (mbApiInterface.Player_GetPlayState() == PlayState.Playing)
                        DisplaySongInfo(sourceFileUrl);
                    break;

                case NotificationType.TrackChanged:
                    DisplaySongInfo(sourceFileUrl);
                    break;

                case NotificationType.NowPlayingArtworkReady:
                    NowPlayingArtworkReady(sourceFileUrl);
                    break;
            }
        }

        private void DisplaySongInfo(String sourceFileUrl)
        {
            String fileUrl = mbApiInterface.NowPlaying_GetFileUrl();
            if (String.IsNullOrWhiteSpace(fileUrl))
                return;

            fileUrl = fileUrl.ToLowerInvariant();
            if (fileUrl == nowPlayingFileUrl)
                return;

            nowPlayingFileUrl = fileUrl;
            nowPlayingHeadline = varHelper.ReplaceNowPlayingVariables(settings.Headline);
            nowPlayingText = varHelper.ReplaceNowPlayingVariables(settings.Text);

            String base64Artwork = mbApiInterface.NowPlaying_GetArtwork();
            if (base64Artwork == null)
                base64Artwork = mbApiInterface.NowPlaying_GetDownloadedArtwork();

            if (String.IsNullOrWhiteSpace(base64Artwork))
                snarl.SongChanged(nowPlayingHeadline, nowPlayingText, defaultIconPath, null);
            else
                snarl.SongChanged(nowPlayingHeadline, nowPlayingText, null, base64Artwork);
        }

        private void NowPlayingArtworkReady(String sourceFileUrl)
        {
            if (sourceFileUrl == null)
                return;
            sourceFileUrl = sourceFileUrl.ToLowerInvariant();

            if (sourceFileUrl != nowPlayingFileUrl)
                return;

            String base64Artwork = mbApiInterface.NowPlaying_GetDownloadedArtwork();
            if (base64Artwork != null)
                snarl.SongChanged(nowPlayingHeadline, nowPlayingText, null, base64Artwork);
        }

        //private String CreateArtworkFileUrl(string base64Artwork)
        //{
        //    if (String.IsNullOrWhiteSpace(base64Artwork))
        //        return null;

        //    var byteArray = Convert.FromBase64String(base64Artwork);
        //    TypeConverter tc = TypeDescriptor.GetConverter(typeof(Bitmap));
        //    Bitmap bitmap = (Bitmap)tc.ConvertFrom(byteArray);

        //    String imgPath = @"C:\Users\tn\AppData\Local\Temp\snarltest.png";
        //    bitmap.Save(imgPath);

        //    return imgPath;

        //    // Bitmap artwork = (Bitmap)bitmapConverter.ConvertFrom(Convert.FromBase64String(base64Text))
        //}

        #region " Not used exported functions "

        // return an array of lyric or artwork provider names this plugin supports
        // the providers will be iterated through one by one and passed to the RetrieveLyrics/ RetrieveArtwork function in order set by the user in the MusicBee Tags(2) preferences screen until a match is found
        public string[] GetProviders()
        {
            return null;
        }

        // return lyrics for the requested artist/title from the requested provider
        // only required if PluginType = LyricsRetrieval
        // return null if no lyrics are found
        public string RetrieveLyrics(string sourceFileUrl, string artist, string trackTitle, string album, bool synchronisedPreferred, string provider)
        {
            return null;
        }

        // return Base64 string representation of the artwork binary data from the requested provider
        // only required if PluginType = ArtworkRetrieval
        // return null if no artwork is found
        public string RetrieveArtwork(string sourceFileUrl, string albumArtist, string album, string provider)
        {
            //Return Convert.ToBase64String(artworkBinaryData)
            return null;
        }

        #endregion

        #region " Storage Plugin "
        // user initiated refresh (eg. pressing F5) - reconnect/ clear cache as appropriate
        public void Refresh()
        {
        }

        // is the server ready
        // you can initially return false and then use MB_SendNotification when the storage is ready (or failed)
        public bool IsReady()
        {
            return false;
        }

        // return a 16x16 bitmap for the storage icon
        public Bitmap GetIcon()
        {
            return new Bitmap(16, 16);
        }

        public bool FolderExists(string path)
        {
            return true;
        }
        
        // return the full path of folders in a folder
        public string[] GetFolders(string path)
        {
            return new string[]{};
        }

        // this function returns an array of files in the specified folder
        // each file is represented as a array of tags - each tag being a KeyValuePair(Of Byte, String), where Byte is a FilePropertyType or MetaDataType enum value and String is the value
        // a tag for FilePropertyType.Url must be included
        // you can initially return null and then use MB_SendNotification when the file data is ready (on receiving the notification MB will call GetFiles(path) again)
        public KeyValuePair<byte, string>[][] GetFiles(string path)
        {
            return null;
        }

        public bool FileExists(string url)
        {
            return true;
        }

        //  each file is represented as a array of tags - each tag being a KeyValuePair(Of Byte, String), where Byte is a FilePropertyType or MetaDataType enum value and String is the value
        // a tag for FilePropertyType.Url must be included
        public KeyValuePair<byte, string>[] GetFile(string url)
        {
            return null;
        }
        
        // return an array of bytes for the raw picture data
        public byte[] GetFileArtwork(string url )
        {
            return null;
        }

        // return an array of playlist identifiers
        // where each playlist identifier is a KeyValuePair(id, name)
        public KeyValuePair<string, string>[] GetPlaylists()
        {
            return null;
        }

        // return an array of files in a playlist - a playlist being identified by the id parameter returned by GetPlaylists()
        // each file is represented as a array of tags - each tag being a KeyValuePair(Of Byte, String), where Byte is a FilePropertyType or MetaDataType enum value and String is the value
        // a tag for FilePropertyType.Url must be included
        public KeyValuePair<byte, string>[][] GetPlaylistFiles(string id)
        {
            return null;
        }

        // return a stream that can read through the raw (undecoded) bytes of a music file
        public System.IO.Stream GetStream(string url)
        {
            return null;
        }

        // return the last error that occurred
        public  Exception GetError()
        {
            return null;
        }

        #endregion
    }
}