# Music Directory Syncer

Music Directory Syncer is a Windows application that allows a directory of music files to be synced to any number of others, according to certain rulesets:

  - Tags being present
  - Tags having certain values
  - File type

Audio files can be converted to a chosen lossy format using ffmpeg if desired to save space, e.g. on a mobile device. This can be done to all files or just losslessly compressed ones. You can also choose to apply ReplayGain volume adjustment if the device you're using doesn't support ReplayGain tags. Once you've made your synced directory, you can use [any directory syncing application] to copy it to a mobile device, for example.

![Create New Sync](/Screenshots/Create%20New%20Sync%20Window.png?raw=true "Create New Sync Window")

Music Directory Syncer will run quietly in the background watching your chosen source directory for changes, and it will update all of your sync directories on-the-fly.

[//]: # (These are reference links used in the body of this note and get stripped out when the markdown processor does its job.)


   [any directory syncing application]: <https://play.google.com/store/apps/details?id=dk.tacit.android.foldersync.lite>
  

