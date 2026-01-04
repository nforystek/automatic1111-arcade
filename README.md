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

Users put in 0.25 cents and get 3 credits.  Each image generated costs 2 credits, and to vote on any one generated image to be in the top voted on list by votes is 1 credit.  Appropriately the font and generating image sound is Transformers(R).  Idle it makes more casino type of sounds periodically, and there is a coin drop 

First, you will need to get [Automatic1111](https://github.com/AUTOMATIC1111/stable-diffusion-webui) setup.  You might want to just start with this article however:  [How to Run Stable Diffusion Locally With a GUI on Windows](https://www.howtogeek.com/832491/how-to-run-stable-diffusion-locally-with-a-gui-on-windows/#how-to-install-stable-diffusion-with-a-gui) as it can be very difficult depending on your system setup.  There are links below for a very similar system to the one I wrote the arcade on.  After you have Automatic1111 able to load and generate a image from text without errors, I recommend also including the following Automatic1111 extensions (and only these extensions): [stable-diffusion-webui-auto-tls-https](https://github.com/papuSpartan/stable-diffusion-webui-auto-tls-https.git) and [stable-diffusion-webui-nsfw-censor](https://github.com/AUTOMATIC1111/stable-diffusion-webui-nsfw-censor.git)

The application is written in Visual Basic 6, Visual C++ 6 and Python 3.10.6 requiring "Microsoft ActiveX Data Objects 2.8 Library" for a Microsoft Access 97 database that it stores everything (obtained with MDAC\_TYP.EXE). The two supporting C++ DLL's I did not write, their sources are included in this repository (and were/are available on the internet which is where I acquired them).  One of which is Redirect.DLL, it allows Visual Basic 6 to open console applications (not visible) and read/write to console apps i.e. programmable "cmd.exe."  The second DLL is inpout32.dll (and the 64bit version inpoutx64.dll) which is an API interface used to read and write to the data IO logics of a parallel port for this application to hardware detect a coin box's switch when a coin is put in it.

IF YOU INTEND TO TRY TO RUN THIS PROGRAM, PLEASE READ THE COMPILE CONDITIONS COMMENTS IN THE SOURCE CODE CAREFULLY AS IT IS POSSIBLE TO LOCK YOURSELF OUT OF THE SYSTEM LEAST THE POWER BUTTON AVAILABLE FOR RECOVERY TO BE USED.

The application runs as a Windows Shell with all restrictions necessary to be a public Kiosk with a keyboard.  In a full production environment (inside a coin operated arcade box housing) auto logon to a password-less administrator account is the only thing that does not setup with the code
and would be required to boot it straight into the shell.  Just ask CoPilot "What are the registry entries used to auto login to a password-less account on Windows 11?" coPilot basically wrote most the Visual Basic 6 registry module and the Txt2Img2Txt.py Python.

A list of products for a keyboarded kiosk: [Amazon.com Automatic1111 Arcade Hardware List](https://www.amazon.com/hz/wishlist/ls/3R5Y55014VKWR?ref_=wl_share)

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
        all time accumulating, but they reset if the machine goes to term elections and
        the term amount is set in the code via constants documented, the auto switching
        is based on usage that can also be set with the term threshold, for instance if
        the usage is lesser then more, the no term will be enabled, when many uses occur
        for high traffic arcades, the term stays defined by days passing, resetting votes.
        Note, a resource file exists mapped out in directories so you can delete mine for
        your own conscious and build it yourself, a plus is it consists of custom sound FX.

####################################################################################################
####################################################################################################
####################################################################################################
```

