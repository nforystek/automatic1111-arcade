# automatic1111-arcade
<p align=middle>
<center>
<table>

<tr><td colspan=2>
<img src="./sd-arcade.png" alt="Stable Diffusion Automatic1111 Arcade" width="200"/>
</td></tr>

<tr><td>
<img src="./ScreenShots/Screenshot (01).png" alt="The startup screen." width="320" height="240"/>
</td><td>
<img src="./ScreenShots/Screenshot (02).png" alt="Main periodic ad screen." width="320" height="240"/>
</td></tr>

<tr><td>
<img src="./ScreenShots/Screenshot (03).png" alt="An empty leaderboard screen." width="320" height="240"/>
</td><td>
<img src="./ScreenShots/Screenshot (04).png" alt="A blank prompt screen." width="320" height="240"/>
</td></tr>

<tr><td>
<img src="./ScreenShots/Screenshot (05).png" alt="No images generated screen." width="320" height="240"/>
</td><td>
<img src="./ScreenShots/Screenshot (06).png" alt="A prompt screen with a prompt entry." width="320" height="240"/>
</td></tr>

<tr><td>
<img src="./ScreenShots/Screenshot (07).png" alt="A generated image screen from the prompt." width="320" height="240"/>
</td><td>
<img src="./ScreenShots/Screenshot (11).png" alt="A gallery screen, allowing you may vote." width="320" height="240"/>
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

An app written in VB6 and C++ for Windows 11 that runs as a legacy Kiosk and intends to be a coin operated Text-to-Image Arcade setup using Stable Diffusion Automatic1111 API (on localhost, without internet nor mouse).

Users put in 0.25 cents and get 3 credits.  Each image generated costs 2 credits, and to vote on any one generated image to be in the top 12 is 1 credit.

The following Automatic1111 extensions are involved/supported:

stable-diffusion-webui-auto-tls-https	https://github.com/papuSpartan/stable-diffusion-webui-auto-tls-https.git

stable-diffusion-webui-nsfw-censor	https://github.com/AUTOMATIC1111/stable-diffusion-webui-nsfw-censor.git

The main application is written in Visual Basic 6, Visual C++ 6, Python 3.10.6 and requires "Microsoft ActiveX Data Objects 2.8 Library" for a Microsoft Access 97 database that it stores everything (you can get with MDAC\_TYP.EXE). The two supporting C++ DLL's I did not write, their sources are included in this repository (and were/are available on the internet which is where I acquired them).  One of which is Redirect.DLL, it allows Visual Basic 6 to open console applications (not visible) and read/write to console apps i.e. programmable "cmd.exe."  The second DLL is inpout32.dll (and the 64bit version inpoutx64.dll) which is an API interface used to read and write to the data IO logics of a parallel port for this application to hardware detect a coin box's switch when a coin is put in it.

IF YOU INTEND TO TRY TO RUN THIS PROGRAM, PLEASE READ THE COMPILE CONDITIONS COMMENTS IN THE SOURCE CODE CAREFULLY AS IT IS POSSIBLE TO LOCK YOURSELF OUT OF THE SYSTEM LEAST THE POWER BUTTON AVAILABLE FOR RECOVERY TO BE USED.

The application runs as a Windows Shell with all restrictions necessary to be a public Kiosk with a keyboard.  In a full production environment (inside a coin operated arcade box housing) auto logon to a password-less administrator account is the only thing that does not setup with the code
and would be required to boot it straight into the shell.  Just ask CoPilot "What are the registry entries used to auto login to a password-less account on Windows 11?" coPilot basically wrote most the Visual Basic 6 registry module and the Txt2Img2Txt.py Python.

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
        credit in the game, the screens do not auto switch unless it's for new generate.
        While credits remain, all screens are available via navigation of the keyboard.

####################################################################################################
####################################################################################################
####################################################################################################
```

Incomplete list of products (missing the monitor and speakers, which it appears much of the customization can come from the cabinet maker) for a keyboarded kiosk (all the buttons for navigation, every least the prompt entry, are touch enabled, or have buttons relating to the key shortcuts).


[LVL23C 54" Tall 2 Player Cabaret Arcade Cabinet Kit](https://www.amazon.com/LVL23C-Player-Cabaret-Cabinet-buttons/dp/B0CJ5ZSZRM/?_encoding=UTF8&pd_rd_w=pCyV7&content-id=amzn1.sym.4efc43db-939e-4a80-abaf-50c6a6b8c631%3Aamzn1.symc.5a16118f-86f0-44cd-8e3e-6c5f82df43d0&pf_rd_p=4efc43db-939e-4a80-abaf-50c6a6b8c631&pf_rd_r=4XQ3WPWS0WVDDBYB5GNH&pd_rd_wg=Gax4g&pd_rd_r=26d6c64d-efc4-4ba3-af94-98fbe466fe34&ref_=pd_hp_d_atf_ci_mcx_mr_ca_hp_atf_d)


[Dell XPS 8960 Desktop - 14th Generation Intel Core i9-14900K Processor, 32GB DDR5 RAM, 2TB SSD, NVIDIA GeForce RTX 4070 12GB GDDR6X Graphics, Windows 11 Pro, Onsite & Migrate Service - Platinum Silver](https://www.amazon.com/Dell-XPS-8960-Desktop-Generation/dp/B0CYT9K6PD/ref=sr_1_15?dib=eyJ2IjoiMSJ9.aZK-Mrzvwu_x1-xOn5CnfaOA8jiDXa3VCN_uqzV28gyBMXzot966hQQIHvrqmYEoCmsCt-ahzZZN9_U1CHoyw38GY8twYLeyNmOvv7nBFhXF8qkh7OEfMSaJ2AT6C9THebuu5QBkUCm3ZK8vNkU-bmplFZxT0-5W_B5uB5BrAuzD0of7hMJVfRiDzVvhumIf4G9eVVRPYhFYLk7Q-TxyGJ24XXY39FOhsY4R_AJ1vKO3pLQRpPTIuwepM2sWuI_YsBXGyB-I6E-ZcmErImguNFUNOP4fDJLucvLQtUwwLq8.VHXg_Qk3nSvaMNHncMixTha9fwNmx8fge-zt5U6jjqg&dib_tag=se&keywords=xps%2B8960%2Bnvidia%2Brtx%2B4060&qid=1767194150&s=electronics&sr=1-15&th=1)


[High-Speed Parallel Port PCI Express Printer I/O Card](https://www.amazon.com/dp/B0BJ7NC3D4?ref_=ppx_hzsearch_conn_dt_b_fed_asin_title_1)


[VHV Arcade Coin Door V2 Sticker Vinyl Waterproof Sticker Decal Car Laptop Wall Window Bumper Sticker 5inch VHV-CRYPTO-STICKERS-4620](https://www.amazon.com/VHV-Arcade-Sticker-Waterproof-VHV-CRYPTO-STICKERS-4620/dp/B09KDJ4JFY/ref=sr_1_19?dib=eyJ2IjoiMSJ9.Ic1BWKBzYipmzDsms9kiE9LmAMVm81Xw-j_1iKG81pYG-8W2RSYbqYQAe22Jg6etQ05yLyc0a6403fSVoCaAhVDvqPAXYQj8GT3XYSjoN51yOv31f-icFNP7QvJoJaQ0htRcqUklCz2vje9HSlqX7KPxx5JPOQw5MqK6VbasSpXJRYHyXoNuQaGRpBh9gu25q1VGp8pHtLBI1ZseA1qgRph4S8xZCWx4uyrnydoZ85aH7cSdeckB2TT_vUCwAz9bezs10qUrL1WFhE5E1IMSRyLcb1fZgCRr5XMG7FdceXs.Ew4JbiSs3pD8fvX8j7Zi0AsKT7cZxnZP842NVM3LAa4&dib_tag=se&keywords=arcade+coin+box&qid=1767194400&sr=8-19)


[New Wave Toys RepliTronics Arcade Change Machine USB-C x2 and USB-A x4 Charging Station, Retro Desktop Office Hub for Multiple Devices, Classic Wood Panel](https://www.amazon.com/dp/B085FXGH38/ref=sspa_dk_detail_5?psc=1&pd_rd_i=B085FXGH38&pd_rd_w=Bvew4&content-id=amzn1.sym.953c7d66-4120-4d22-a777-f19dbfa69309&pf_rd_p=953c7d66-4120-4d22-a777-f19dbfa69309&pf_rd_r=P2M78K3QWKWNQ3S9PJAS&pd_rd_wg=CMiAN&pd_rd_r=268ed41c-6131-478e-8c40-374f519bfdc7&sp_csd=d2lkZ2V0TmFtZT1zcF9kZXRhaWwy)


