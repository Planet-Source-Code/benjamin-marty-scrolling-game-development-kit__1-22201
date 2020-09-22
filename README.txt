The Scrolling Game Development Kit
ReadMe File

March 20, 2001

The Scrolling Game Development Kit (GameDev) is designed and developed by
Benjamin Marty (BlueMonkMN@email.com).  GameDev and its components are
distributed under the terms of the GNU General Public License (GPL).
See COPYING.txt for more information.

This file is included with each package related to The Scrolling Game
Development kit (older versions may exist in files that haven't needed updates).
Depending on which package(s) you download, different sections may apply.
It covers binaries and source code for GameDev, BMDXCtls
and ScrHost.  To access the main distribution point for all pieces related to
GameDev, visit http://www.sourceforge.net/projects/gamedev/ or visit the GameDev
homepage at http://gamedev.sourceforge.net/

Release Notes
=============

To see what's new in this 1.2 release of GameDev, check out the "What's New"
topic in the help file.

Changes specific to version 1.2.1 are:
* Tile Animations are now reset before opening a map so if you have different
  tile animations that need to be synchronized with each other, that will work
  now.  (Multiple tiles using the same animation always worked fine, BTW)
* Now you can import a 64x64 pixel tile from a 64x64 pixel bitmap (rounding
  error fixed which displayed "bitmap not wide enough" error before).
* In some cases, a sprite template would forget which solidity definition it
  was using when being loaded from a .MAP file.  That has been fixed.
* Now you can play video clips that do not include sound.
* The names of the collision classes (seen in the last tab of the sprites and
  paths dialog) are now stored during an XML export so you don't lose them when
  playing a game from the dev environment or while exporting to XML.
Note: All the above bugs except the last one went unnoticed in versions 1.0
      through 1.2.  The last one was introduced in version 1.2.

There is still no documentation for the scripting support or COM support
exposed by GameDev except the last page of the Tutorial.  Anyone interested in
helping out with this, please contact me.

The .vbs template files and the sample project (included since prior release)
are good examples of how to use scripting to control GameDev at runtime *and*
at design time (for instance, to assist with editing maps).  The DngnEdit.vbs
script allows you to simply press the "f" key to automatically fill the screen
with the appropriate tiles to cover up the player when he's behind a wall and
to fill out the "front face" of walls not completed by the "WallTop" tilematch.
The sample project demonstrates how one might implement an isometric-like view
in GameDev.

GameDev uses ScrHost.dll specifically created for the purpose of hosting
ActiveX script in GameDev.  (At the time, all I had was VB 5 and the script
hosting ActiveX control did not seem to be working in VB5.)

This program also uses BMDXCtls.dll version 1.2 which is being released with
the initial release of GameDev as one of its component modules.  Previous
versions of BMDXCtls have been completed and released (closed source) earlier
as an independent library component available to other developers.  However
the intent all along (since 1998) was to include this component in a product
such as this.  For more information about BMDXCtls see the sections below.

The main documentation for GameDev is distributed with the installation package,
but is also available on the web at this projects home site:
http://gamedev.sourceforge.net/
Also, refer to this site for up to date information, bug reports, other tech
support, etc.  The project is open source and hosted by SourceForge.net at
http://www.sourceforge.net/projects/gamedev/


Requirements
============

Hardware:

Most any system capable of running Windows 95 should be capable of running
GameDev.
Beyond this, GameDev does require a video card with DirectX drivers that can
support 16-bit, 24-bit or 32-bit display depths (these are the only video modes
supported by BMDXCtls, on which GameDev depends).  It uses these color depths at
a resolution of 640x480.
Performance will be greatly improved with the quality of DirectDraw drivers and
Video Memory (2-D).  BMDXCtls depends greatly on the speed that the video card
can transfer graphics from an off-screen surface to the display.  More video
memory means more graphics stored in the video card, which in turn means
speedier transfer of graphics.

Software:

GameDev installs BMDXCtls to interface with DirectX.  This component depends
on DirectX 5.0 or later.  GameDev does not include installation of DirectX.
Games that use multimedia clips should be aware of the codecs required to play
back these clips -- they may not reside on all other systems where the game
might be played.  GameDev uses quartz.dll (ActiveMovie) for media playback.
XML support requires Microsoft's XML COM object installed (I think) with IE5

GameDev also installs ScrHost.dll to host VBScript code.  This requires that
Windows Script Components be installed.  If you do not already have this
(it should be included with IE4.01 and later) it can be obtained from
http://msdn.microsoft.com/scripting/
GameDev does not include installation of Windows Script Components.


About BMDXCtls License
======================

As of version 1.2 BMDXCtls is becoming open source in correspondence with the
GPL.  BMDXCtls normally displays a dialog when the full screen display is
opened.  Previous versions of BMDXCtls required software developers to contact
the author of BMDXCtls in order to get a code to turn off this dialog box.
Now that the project is open source, this is easily circumventable.  It's
possible to recompile the code without the dialog box or to determine the
appropriate code to turn the dialog off.  All I ask is that credit be given
where credit is due.  I think the "ValidateLicense" method in BMDXCtls is still
appropriate (this is the call to turn off the dialog box).  It's useful to make
sure that someone who installs GameDev binaries doesn't see this handy component
and use it for themselves without giving credit or even seeing license
information.  Therefore I'm leaving the "ValidateLicense" method in BMDXCtls.

My suggestion is we all stick with one binary and if you really want to use this
in your own project and turn off the dialog, you're free to step through the
source code in a debugger and get the appropriate license string to put into
your project.  Note that the license string is based on the EXE name, so your
project must be compiled and run outside the development environment, otherwise
BMDXCtls will pick up the EXE name of the IDE.

Also, as of BMDXCtls version 1.2, BMDXCtls is part of the GameDev distribution.
I'm not trying to maintain it as its own product any more, but nothing was lost.
Some funcionality was added (and documented in the BMDXCtls help file).

About ScrHost
=============

This was created for the sole purpose of supporting ActiveX scripting in
GameDev.  I'm releasing it with GameDev under GPL.  It's lightly documented via
type library help strings.  (Reference this component in Visual Basic and bring
up the object viewer with F2.  Short help strings can be seen associated with
each member of this library.)

Distributing Games / Playing Games
==================================

GameDev.exe will offer to register the GDP file type if it's not already
registered when you run it.  Once this occurs, you can right click on any GDP
file and click "Play" to play the game, assuming it's complete and set up to
play.  The command line to play a game is:
GameDev.exe <game.gdp> /p

If you answered "No" when asked if you would like to register the "GDP" file
extension with GameDev, and have changed your mind, or if you ever run into
problems with window positions or default directory settings (possibly even
errors related to these that I have not yet encountered in testing) you can
delete GameDev's registry settings at:
HKEY_CURRENT_USER\Software\VB and VBA Program Settings\GameDev
This does not delete the GDP file association.  To do this go to the Folder
Options dialog in Windows Explorer, select the File Types tab, locate the entry
for "GameDev Project", select it and click Remove.

There are currently no protections to protect a game from being loaded into
gamedev and edited once it's complete.  GameDev is like a "browser" type
program which lets you "view source".

As of GameDev version 1.2, there is a new "Make Shortcut" command that you can
use to assist in the process of creating a shortcut to start-up a project in
playback mode.

Recommendations for distributing a game:
* Keep all files relavent to your game in a single directory.  All paths stored
  in the GDP file are relative to the GDP file.
* Zip the directory up and send it / post it wherever.
* Post a comment about the availability of the new game at the project's home
  site.
* Include a readme with your game indicating where to get GameDev in case
  the downloader doesn't realize it's a GameDev project.
* See the GPL (COPYING.txt) for terms of distributing GameDev with your
  own distribution.  Basically, the GameDev installation package can be
  distributed without alteration, but if you need to make modifications to
  include your own piece, you must release source code and make clear
  indications of the changes.
* If you want to create an icon to run the project on the target system, you
  will need to create some sort of installer that can locate GameDev on the
  target system and use that path to create a shortcut to the GDP file,
  including the "/p" command line switch.  The path to GameDev can be found
  in the registry at:
  HKEY_CLASSES_ROOT\CLSID\{1DDB0F14-53D0-4EC4-9D13-CB2AA95598FB}\LocalServer32
  or:
  HKEY_CLASSES_ROOT\TypeLib\{D0BC0A98-5AF6-4B55-9DCB-F6ABB4895D9D}\1.2\0\win32
  (The second depends on a specific version, 1.2 in this case)


Source Code / Build Information
===============================

The GameDev installation installs a completely functioning system of binary
files.  This section applies to rebuilding source code if you have a task
that requires more than the basic binary setup (for instance debugging or
altering code).

GameDev itself is a Visual Basic program developed under Microsoft Visual
Basic 5.0 and 6.0 (finally built with 6.0).  Before building this project you
must first build the two component DLLs (which will register them in the
process):

(Before beginning, you may want to un-register and delete or rename any
existing copies of BMDXCtls.dll and ScrHost.dll you already have installed so
they do not interfere with the newly registered versions being compiled here).

BMDXCtls: This is a C++ project which utilizes and exposes features of
          Microsoft DirectX to Visual Basic.  It was developed in Microsoft
          Visual C++ 5.0 and 6.0 and finally compiled in MSVC 6.0.  To compile
          the project load BMDXCtls.dsw into Microsot Visual C++ 6.0, select
          the "Win32 Release MinDependency" target and build.  This should
          generate and register a DLL file in a directory called
          "ReleaseMinDependency" under the BMDXCtls project directory.  You
          may need to copy the help file from BMDXCtls into this subdirectory
          in order to be able to access online help for BMDXCtls from Visual
          Basic.  BMDXCtls.rtf and BMDXCtls.hpj are the source code files for
          the help file.  HCW.EXE (included with Visual Studio) is used to
          maintain and compile the help file.

ScrHost:  This is also a C++ project developed in Microsoft Visual C++ 6.0.  It
          exposes ActiveX script hosting ability to a Visual Basic project
          (as existing support did not seem to work in Visual Basic 5.0).  To
          build it, load ScrHost.dsw into Microsoft Visual C++ 6.0, select the
          "Win32 Release MinDependency" targer and build. This should generate
          and register ScrHost.dll in the ReleaseMinDependency subdirectory.
          There is no online help for this DLL -- only type library help.

GameDev:  Now that the component DLL's are registered, GameDev.vbp can be
          loaded into the Microsoft Visual Basic 6.0 development environment.
          Make sure you have a file called "GameDev.cmp" in the project
          directory.  This should be about the size of the GameDev.EXE file.
          It is the "Version Compatible Component" used by Visual Basic at
          compile time to ensure that the COM interfaces in the new exe file
          are compatible with those in previously compiled versions.
          The project should load now, and be able to reference the two DLL
          files build above as well as the CMP file.  All you need to do is
          select Make from the file menu to build this project (it's rather
          slow to build -- big project I guess).
          If you attempt to compile this and find that compatibility must be
          broken (an interface changed)... don't do it?  Handling all the
          intricacies of breaking the interface and redistributing an open
          source project and binary is beyond the scope of this paragraph.
          The help file is GameDev.chm.  All the source code for this is
          included in the Help directory under GameDev.  This is an HTML Help
          Project compiled with the HTML Help workshop available with Visual
          Studio 6.0 Service Pack 4.  GameDev.hhp is the top level help project
          file.  All images referenced by the help are in the Images
          subdirectory.
          The Res subdirectory of GameDev contains images used while creating
          the project.  They are not required to compile the project because
          they are part of the binary data of the project (frx files).  But
          should they need to be updated and re-inserted into the project,
          these are the source images.
          Finally, the PDM file is the file used by the package and deployment
          wizard to remember how to create a setup package for GameDev.  You
          will need to change the paths in this file if you want to use the
          same file... it contains hard coded full paths to other components.
          I suggest assembling the package into the "Package" subdirectory of
          GameDev (which may or may not exist, depending on if the "Package"
          directory ends up being part of the source code package).  Note the
          files referenced in the PDM file and the installed locations.  These
          files are the list of files distributed and installed with the
          official binary installation package.


Anything Missing?
=================

If there's anything missing here, visit http://gamedev.sourceforge.net/ and
http://www.sourceforge.net/projects/gamedev/ for the latest.  Post your own
comments or questions there.