# ucAniGifEx v1.1.3
ucAniGifEx Animated Gif Control

![ss](https://github.com/user-attachments/assets/ad8fcf31-6807-4193-9550-4b00f8b19a91)

**ucAniGifExTest.twinproj** - Project with test form shown in the picture above. Requires twinBASIC Beta 734 or newer since the gif file list box gives paths with a backslash in prior versions.

**AniGifEx.twinproj** - ActiveX control project to build .ocx versions for VB6/VBA/etc

See [Releases](https://github.com/fafalone/ucAniGifEx/releases) for binaries.

**PROJECT UPDATED:**\
(v1.1.3, 15 Apr 2025) DisplayGifFromURL, Click/DblClick events, ResetPlayback

(v1.0.1, 09 Apr 2025) Quick fix for various looping issues, see Readme below.

```
/********************************************************************************
    ucAniGifEx v1.1.3
    by Jon Johnson, ported from Windows SDK WicAnimatedGif example, with
         corrections, additional features, and conversion to control.
    (c) 2025. Distributed under the MIT License.
    
    This is a followup to my ucAniGif control, which while it had very 
    simple code, there were a number of flaws. Some were mistakes that could
    be fixed, but others were limitations of the underlying GdiPlus code or
    impractical to fix in the shell interface (I already was resorting to
    patching assembly instructions in memory to fix the slow playback bug!).
    
    Many gifs did not play right, or at all. And very limited extras.
    
    While the code here is significantly more complex, this control supports 
    every gif I've found, but it's just as simple to use from the outside 
    while adding some basic enhancements.
        
    Features:
      -Multiple options to display a gif: 
         StartupFile property - Loads and plays a gif file from disk on load
         DisplayGifFromFile - Loads and plays a gif file from disk on call
         DisplayGifFromResource - Loads and plays a gif file from the project
              resources. Specify the ID and group; LoadResData is used.
         DisplayGifFromMemory - Play direct from a byte array containing a 
              complete gif file.
         DisplayGifFromURL - Play from an http:, https:, or ftp: resource.
      
      -Paused property can suspend/resume playback, StopPlayback, 
          ResetPlayback commands.
      
      -Will be transparent to the control BackColor; it's recommended you set
       this before playing, i.e. ucAniGifEx1.BackColor = Me.BackColor, then play.
      
      -Gif is scaled to the control size by default.
      
      -PreserveAspectRatio option, with centering in the frame
      
      -SizeToFit - Resizes the control to the gif size. Scaling is disabled.
      
      -LoopControl - Override the gifs default to force infinite loops or only
          play once.
      
      -LoopEnd, PlaybackEnd, and FrameStep events. Note: FrameStep is disabled
          by default, to receive it, set EnablePerFrameEvents to True.
      
      -AdvanceByOneFrame - If paused, advance by a single frame without resuming.
      
      -FrameCount, FrameIndex, ImageWidth, ImageHeight properties.
    
      -Supports DPI Aware apps
      
    Requirements:
      Windows 7 or newer
      
    Known issues:
      twinBASIC currently displays a continuable exception error in the IDE. This
      does not seem to be impacting the control. A bug report has been filed.
      
    Change log: 
      (15 Apr 2025, v1.1.3)
          -Added DisplayGifFromURL to play from http: https: and ftp:
          -Added Click and DblClick events
          -Added ResetPlayback method to restart animation from first frame.
          

      (09 Apr 2025, v1.0.1)
          -To match modern browsers, gifs without looping info will only play once.
          -Bug fix: Loop control override not properly implemented.
          -Bug fix: Missing ByVals in loop metadata lookup caused all gifs to loop 
                infinitely; could possibly cause crash.
       
      (08 Apr 2025, v1.0) - Initial release.
      
    Thanks
      Special thanks to bahbahbah (whose code I wound up using), jpbro, and 
      The trick for helping me sort through some scaling issues.
      And wqweto for helping me out with some C++ issues to get the original
      SDK example working so I could compare.
*********************************************************************************/
```

### How to install and select for use in VBA

#### From binaries

There's already-compiled binary builds available under the 'Releases' link on the right side column of this page, [or click here](https://github.com/fafalone/ucAniGifEx/releases) to go directly to the Releases page. The  zip contains a win32 folder and a win64 folder, you need to use the one that matches Office for 32bit or 64bit. Once you've extracted the one you need, if it's 32bit, drag/drop `AniGifEx.ocx` in Explorer onto the `regsvr32.exe` in the Windows\SysWOW64 folder. For 64bit, drop on the one in Windows\System32 (unless you're using 32bit version of Windows, then just that too).

> [!TIP]
> If you don't know whether you have 32bit or 64bit Office, go to File->Account then click 'About Excel/Access/etc'


#### From source

As noted above, this project was written in twinBASIC. If you don't have it already you can download it from [here](https://github.com/twinbasic/twinbasic/releases). Just download and extract, there's no installer. The Community Edition is free, and will build both 32bit and 64bit OCXs-- the only limitation being that the 64bit OCX will display a splash screen for tB when it loads. The OCX in the Releases section here will not as I've got the Pro version.\
You need the AniGifEx.twinproj file from the source code files of this repository. That contains the entire project, the other files are for browsing the source code in your browser or the ucAniGifExTest.twinproj file is a demo of using the control within twinBASIC. Use the 'Browse' option to open the .twinproj file, in the bottom left of the opening dialog when you launch twinBASIC. After that, look in the toolbar for a dropdown list that says 'win32'... that means it will compile as 32bit. Select win64 from the dropdown to compile as 64bit. After that, click on File->Build. It will create and automatically register the OCX, you're all set to open Office and start using it

#### Finding it in VBA

Once you've done one of the options above, `AniGifEx` should be available in the Tools->Additional controls dialog when you're editing a UserForm in Excel VBA, or 'ActiveX Controls' in the Access form designer-- the menu that pops up from the dropdown button on the righthand side of the built-in controls box.
