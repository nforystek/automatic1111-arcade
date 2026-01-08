# automatic1111-arcade
<p align=middle>
<center>
<table>

<tr><td colspan=2 align=middle><center>
<img src="./sd-arcade.png" alt="Stable Diffusion Automatic1111 Arcade (digital mock)" width="200"/>
</center>
</td></tr>

<tr><td>
<img src="./ScreenShots/Screenshot (01).png" alt="The startup screen." width="320" height="240"/>
</td><td>
<img src="./ScreenShots/Screenshot (02).png" alt="Main periodic ad screen." width="320" height="240"/>
</td></tr>

<tr><td>
<img src="./ScreenShots/Screenshot (03).png" alt="An empty leaderboard screen." width="320" height="240"/>
</td><td>
<img src="./ScreenShots/Screenshot (04).png" alt="No images generated screen." width="320" height="240"/>
</td></tr>

<tr><td>
<img src="./ScreenShots/Screenshot (05).png" alt="A blank prompt screen." width="320" height="240"/>
</td><td>
<img src="./ScreenShots/Screenshot (06).png" alt="A prompt screen with text entered." width="320" height="240"/>
</td></tr>

<tr><td>
<img src="./ScreenShots/Screenshot (07).png" alt="A generated image screen from the prompt." width="320" height="240"/>
</td><td>
<img src="./ScreenShots/Screenshot (08).png" alt="A gallery screen, allowing you may vote." width="320" height="240"/>
</td></tr>

<tr><td>
<img src="./ScreenShots/Screenshot (09).png" alt="A screenshot after of a voted on image." width="320" height="240"/>
</td><td>
<img src="./ScreenShots/Screenshot (10).png" alt="A leaderboard with some images in it." width="320" height="240"/>
</td></tr>

</table>
</p>
</center>


Preface:

A legacy Kiosk and intends to be a coin operated Text-to-Image Arcade setup using Stable Diffusion Automatic1111 API (on localhost, without internet nor mouse).

Users put in 0.25 cents and get 3 credits.  Each image generated costs 2 credits, and to vote on any one generated image to be in the top voted on list by votes is 1 credit.  Appropriately the font and generating image sound is Transformers(R).  Idle it makes more casino type of sounds periodically, and there is a coin drop sound.

First, you will need to get [Automatic1111](https://github.com/AUTOMATIC1111/stable-diffusion-webui) setup.  You might want to just start with this article however:  [How to Run Stable Diffusion Locally With a GUI on Windows](https://www.howtogeek.com/832491/how-to-run-stable-diffusion-locally-with-a-gui-on-windows/#how-to-install-stable-diffusion-with-a-gui) as it can be very difficult depending on your system setup.  There are links below for a very similar system to the one I wrote the arcade on.  After you have Automatic1111 able to load and generate a image from text without errors, I recommend also including the following Automatic1111 extensions (and only these extensions): [stable-diffusion-webui-auto-tls-https](https://github.com/papuSpartan/stable-diffusion-webui-auto-tls-https.git) and [stable-diffusion-webui-nsfw-censor](https://github.com/AUTOMATIC1111/stable-diffusion-webui-nsfw-censor.git)

The application is written in Visual Basic 6, Visual C++ 6 and Python 3.10.6 requiring "Microsoft ActiveX Data Objects 2.8 Library" for a Microsoft Access 97 database that it stores everything (obtained with MDAC\_TYP.EXE). The two supporting C++ DLL's I did not write, their sources are included in this repository (and were/are available on the internet which is where I acquired them).  One of which is Redirect.DLL, it allows Visual Basic 6 to open console applications (not visible) and read/write to console apps i.e. programmable "cmd.exe."  The second DLL is inpout32.dll (and the 64bit version inpoutx64.dll) which is an API interface used to read and write to the data IO logics of a parallel port for this application to hardware detect a coin box's switch when a coin is put in it.

IF YOU INTEND TO TRY TO RUN THIS PROGRAM, PLEASE READ THE COMPILE CONDITIONS COMMENTS IN THE SOURCE CODE CAREFULLY AS IT IS POSSIBLE TO LOCK YOURSELF OUT OF THE SYSTEM LEAST THE POWER BUTTON AVAILABLE FOR RECOVERY TO BE USED.

The application runs as a Windows Shell with all restrictions necessary to be a public Kiosk housed inside a coin operated arcade box with just a keyboard for the user input. I made an [Amazon.com Automatic1111 Arcade Hardware List](https://www.amazon.com/hz/wishlist/ls/3R5Y55014VKWR?ref_=wl_share) list of product that I anticipate I would be buying for the full build.


The flow of the screens:

```text
####################################################################################################
################################################## Idle mode #######################################
####################################################################################################


>------->  Self Advertisement Screen  -time-out->  Leaderboard Screen  -time-out-1->  Show 1st Place
^---------------------------------------------------------------<      -time-out-2->  Show 2nd Place
>---> -time-out-> Random Casino idle Machine Sounds---V         |      - "  - "  +->  etc 12th Place  
^-----------------<-----------------------------------<         ^------<-----------<  <------------V

####################################################################################################
############################################## Insert 1x0.25 coin ##################################
####################################################################################################

Coin Drop into a Machine Sound ---> Text Prompt Screen + 3 Credits Added. ---> -click-generate-> --V
V-------------------------------< Transformers Changing Sound <------------------------------------< 
Single Image View Screen >------> Election Gallery Vote/Views >-----> -alt-key-> Cast Vote For Image
                                  V                         V
         Last Image <-left-arrow- <                         > -right-arrow-> Next Image

####################################################################################################
################################################ Notes #############################################
####################################################################################################


        With/out credits, Last and Next buttons (Page Up an^d Page Down) switch between
        the advertisement screen, leader board screen, and F1 to F12 keys can be used to
        see what image is in 1st to 12th place respectively to the F key pressed. These
        screens will begin to auto switch if the keyboard remains idle.  When there is
        credit in the game, the screens do not auto switch unless it's done generating.
        While credits remain, all screens are available via navigation of the keyboard.
        There is two modes of election periods, one is no term, and votes are based on
        all time accumulating, but they taper if the machine goes to term elections and
        the term amount is set in the code via constants documented, the auto switching
        is based on usage that can also be set with the term threshold, for instance if
        the usage is lesser then more, the no term will be enabled, when many uses occur
        for high traffic arcades, the term stays defined by days passing, for new elect.
        Note, a resource file exists mapped out in directories with some binary files to
        make it a bit easier in starting it up, and it also consists of custom sound FX.

####################################################################################################
####################################################################################################
####################################################################################################
```

Step by step for preparing a Dell XPS with Windows 11 (in example) to run this legacy Kiosk Template application.

1. Crate the Kiosk admin digital card. Using a digital card or flash drive, (which ever the system interface accepts),

  format the media using FAT32. Open Notepad (or a text editor) and create a single file on the new media drive

  by the name of "KIOSK" no file extension. Inside this file, using a text based editor put the following line:

  "6E98DE51-5380-D7AC-D780-5351DE986E6E"

  No quotes, tabs, new lines or spaces, as the single only first line of text in the file KIOSK.

  Remove the media and hold on to it for later. It will become a valuable part of administration.

 

2. Taskbar Search For "Control Panel" and open by clicking it, then find and click "Change User Account Control settings"

  using "Control Panel" search. Set "Choose when to be notified about changes to your computer" to "Never notify"

3. Taskbar Search For "gpedit.msc" -> Run it by clicking it, the group policy editor (gpedit) MMC will open, find the path

  "\Local Computer Policy\Computer Configuration\Windows Settings\Security Settings\Local Policies\Security Options"

  Under Security Options, find "User Account Control: Run all administrators in Admin Approval Mode" set it to "Disabled"

4. Create a suitable location to house the executables, such as C:\Txt2ImgKiosk

  Copy Txt2ImgKiosk.exe and ShellSwap.exe to the chosen location C:\Txt2ImgKiosk

5. Select "C:\Txt2ImgKiosk\Txt2ImgKiosk.exe" in file explorer and right click

  it, then click the Properties menu, and click the Security tab.

 

  Use the "Edit" button to add users or change their permissions to ensure the

  following users and their allowed permissions listed here are set for the file:

    Authenticated Users: Allow: Full Control, Modify, Read & execute, Read, Write

    Local account: Allow: Full Control, Modify, Read & execute, Read, Write

    SYSTEM: Allow: Full Control, Modify, Read & execute, Read, Write

    Administrators: Allow: Full Control, Modify, Read & execute, Read, Write

    Local account: Allow: Read & execute & Read

  To add a user after you click the "Edit" button click the "Add..." button then click the "Advanced..." button

  and then click the "Find Now" button. A list of user accounts will appear at the bottom of the window, choose

  the account you need to add by highlighting it, and then click the "OK" button, and the "OK" button again.

  Repeat these steps for any user not seen in the permission list by clicking the "Add..." button again. When

  you highlight a user, the permissions set for that user show up in a list below, ensure the users match the

  above permissions for "Allow" and leave all the "Deny" unchecked, and non mentioned permissions unchecked too.

  When finished click the "OK" button to return back to the file properties window.

 

  Click the "Compatibility" tab. Click the "Change Settings for all Users" button, ensure all

  options are unchecked, click OK, and again ensure all options are unchecked and click OK.

7. Create a new local user account

  Open Settings:

  Press Windows + I to open Settings.

  Go to Accounts:

  Accounts → Other users.

  Add a new user:

  Click Add account under Other users.

  In the dialog, choose I don’t have this person’s sign-in information.

  Then choose Add a user without a Microsoft account.

  Enter a username and password (don’t skip the password to use auto logon—it’s required), plus security questions.

  Click Next to create the account.

8. Make that user is an administrator

  Stay in Accounts → Other users.

  Select the new account:

  Click the account you just created.

  Change account type:

  Click Change account type.

  In Account type, change from Standard User to Administrator.

  Click OK.

  Now that user is an admin.

9. Enable automatic logon for that account (simplest method: netplwiz)

  Open Run dialog:

  Press Windows + R, type:

  netplwiz

  and press Enter.

  Select the user:

  In the User Accounts window:

  Select the user account you want to auto log in with (the admin account you just created).

  Disable password requirement:

  Uncheck Users must enter a user name and password to use this computer.

  Click Apply.

  Enter credentials:

  A dialog appears asking for the user name and password for automatic logon:

  Confirm the username is correct.

  Enter the password for that account twice.

  Click OK.

  Confirm and close:

  Click OK again to close the User Accounts window.

10. Run "C:\Txt2ImgKiosk\Txt2ImgKiosk.exe" by double clicking it, this will finalize the Kiosk running system.

   Note, the digital card made earlier is only a temporary pause in the Kiosk mode when inserted, the Kiosk

   remains still running, remove the card and the Kiosk returns to normal operation.

