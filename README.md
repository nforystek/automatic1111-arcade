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

Incomplete list of products (missing the monitor and speakers, which it appears much of the customization can come from the cabinet maker) for a keyboarded kiosk (all the buttons for navigation, every least the prompt entry, are touch enabled, or have buttons relating to the key shortcuts).


[LVL23C 54" Tall 2 Player Cabaret Arcade Cabinet Kit](https://www.amazon.com/LVL23C-Player-Cabaret-Cabinet-buttons/dp/B0CJ5ZSZRM/?_encoding=UTF8&pd_rd_w=pCyV7&content-id=amzn1.sym.4efc43db-939e-4a80-abaf-50c6a6b8c631%3Aamzn1.symc.5a16118f-86f0-44cd-8e3e-6c5f82df43d0&pf_rd_p=4efc43db-939e-4a80-abaf-50c6a6b8c631&pf_rd_r=4XQ3WPWS0WVDDBYB5GNH&pd_rd_wg=Gax4g&pd_rd_r=26d6c64d-efc4-4ba3-af94-98fbe466fe34&ref_=pd_hp_d_atf_ci_mcx_mr_ca_hp_atf_d)

[Dell XPS 8960 Desktop - 14th Generation Intel Core i9-14900K Processor, 32GB DDR5 RAM, 2TB SSD, NVIDIA GeForce RTX 4070 12GB GDDR6X Graphics, Windows 11 Pro, Onsite & Migrate Service - Platinum Silver](https://www.amazon.com/Dell-XPS-8960-Desktop-Generation/dp/B0CYT9K6PD/ref=sr_1_15?dib=eyJ2IjoiMSJ9.aZK-Mrzvwu_x1-xOn5CnfaOA8jiDXa3VCN_uqzV28gyBMXzot966hQQIHvrqmYEoCmsCt-ahzZZN9_U1CHoyw38GY8twYLeyNmOvv7nBFhXF8qkh7OEfMSaJ2AT6C9THebuu5QBkUCm3ZK8vNkU-bmplFZxT0-5W_B5uB5BrAuzD0of7hMJVfRiDzVvhumIf4G9eVVRPYhFYLk7Q-TxyGJ24XXY39FOhsY4R_AJ1vKO3pLQRpPTIuwepM2sWuI_YsBXGyB-I6E-ZcmErImguNFUNOP4fDJLucvLQtUwwLq8.VHXg_Qk3nSvaMNHncMixTha9fwNmx8fge-zt5U6jjqg&dib_tag=se&keywords=xps%2B8960%2Bnvidia%2Brtx%2B4060&qid=1767194150&s=electronics&sr=1-15&th=1)

[High-Speed Parallel Port PCI Express Printer I/O Card](https://www.amazon.com/dp/B0BJ7NC3D4?ref_=ppx_hzsearch_conn_dt_b_fed_asin_title_1)

[RS232 Parallel Serial Port 2 Rows DB25 25 Pin D Sub Male/Female Connector + Plastic Assemble Shell Cover VAG Adapter, Pack of 10](https://www.amazon.com/dp/B07FS6VSVR?ref_=ppx_hzsearch_conn_dt_b_fed_asin_title_2)


[20 Gauge 2 Conductor PVC Tinned Copper Electrical Wires, 32.8FT/10M Red and Black Double Parallel Wire,0.5mm²，for DIY Projects, Home Electrical, Power Wiring,Car Speaker Wires, LED Wires.](https://www.amazon.com/Conductor-Electrical-Parallel-0-5mm%C2%B2%EF%BC%8C-Projects/dp/B0FH4ZWCJ2/ref=sr_1_2_sspa?crid=KCV51U4F9I8J&dib=eyJ2IjoiMSJ9.4gMNwI72qTHHlRhswEaJr7-UcHK6EtxlD0fHQx0tJs17BKJMEC6CNDA2YXF3Io8l99fRQvhkZ6ONuOmd9sZxpQBD-Ur94KJ7uAaJ8f2Xb2gYFXWsp-BldAE6hF91b9fORsv5p8ekiuO6uVgWiMWmHsVOmSiaRnwyXYmZJMoKTtSd6z_K9bEN6hjWaxtQZnVqCE71gVcn8359lg8LCy5X4yd8pnNf-8n2NyRNzS2UAno.OZ8CWap5gnq3S2wqf9ZazwQgnD3OR4OVHteig6Mtnrs&dib_tag=se&keywords=speaker%2Bwire&qid=1767410925&sprefix=speaker%2Bwir%2Caps%2C241&sr=8-2-spons&sp_csd=d2lkZ2V0TmFtZT1zcF9hdGY&th=1)


[19 Inch Arcade Game LED Monitor, Jamma Monitor MAME and Cocktail Game cabinets](https://www.amazon.com/Arcade-Monitor-Cocktail-cabinets-industrial/dp/B004QJ6EN8/ref=sr_1_3?crid=3NDQA3RTK72JS&dib=eyJ2IjoiMSJ9.Q9Gs6t5QeF0BdNoTI075wBZpz5KMuwUiDSPR6gpuFlrrReLuFHaqHH5XQAG9fo-DS0XPl0nXVuvepDmKa8KSNxzL04IbvpJwyBjVBGBADmBjQ8K6y7Gc156da-aeXwyDgZ9ONAHLbEF9IG78Wy8eAETstjWCklQQ3gOl80FoGxZZc2TxwrrwMFYgJIdSgZmMbLJCGnxclLPcXclCDXVG1ImtGaAx9DVXNp9i7sE36oM.BZR2p1xvDw9VxkEildNBcJPF0Klhjov-9QxG47UT-5Y&dib_tag=se&keywords=19%22+Square+LCD+Monitor+for+arcade&qid=1767410232&sprefix=19+square+lcd+monitor+for+arcad%2Caps%2C180&sr=8-3)

[Arcade Game Coin Door with 1 Mechanical Coin selector for US 25 Cent, Jamma, Mame and More!](https://www.amazon.com/Arcade-Game-mechanical-selector-Jamma/dp/B00F1YQJJ6/ref=sr_1_2?crid=26MO1ELFBDGD8&dib=eyJ2IjoiMSJ9.MuRB97Zip8KwCg33OF4KqdLmAMVm81Xw-j_1iKG81pabcOz6ZgBVhrJwZwa_6ynnpC8ICyl8Julx1QQyeMjQbg3dKYR-vzj_iX1-_nErOfIWsfoEk7lJi_T75_gKiBEPUCW6m6Zo1IN6t4xeY1vWCMEeQ_o0BUo3upDFKJGJehn1l_Yxh1UrPW4qxvycVFTd-0t_LlftvsiNJOd-QxSVot0nF_HRqCk0Ncqdz6NemXz29PZFtJGXZRqzcOMvRqNocHMB8dtSyfuLkTRllDAmhxLe_rj136jXvS3AFwCrWv0.g5sdLsYh9miX81uJZELdDx_Zc7LPCu1YVPe92tGoSu4&dib_tag=se&keywords=arcade+coin+box%2C+NC+or+NO+switch&qid=1767410395&sprefix=arcade+coin+box%2C+nc+or+no+switc%2Caps%2C187&sr=8-2)


[Poyiccot 3.5mm Speaker Wire Adapter, 3.5mm 4 Pole Stereo TRRS Male Jack to AV 4 Screw Terminal Block Balun Connectors Cable 30cm](https://www.amazon.com/Poyiccot-Plated-4-Screw-Terminal-Connectors/dp/B07PQTKMM7/ref=sr_1_3?crid=17DJWD64293VS&dib=eyJ2IjoiMSJ9.x6u2IJ4-Z4m0wDQlI3OdTT4x-nPGA5ACBUVJzymK47E2AXusgpBiXQVd6DZPs3i1DtzWTrJFL36YtW9lIuf41g3-Ek80tEItOVqCMTdZs0TYQwksDeFLn8_Z3_zbtWtvSeyzihQat6OJGt5iOos7DDG0yg8HlL8qU3Cn6zfKzUCrlk5fUVhZdbdC-g8iqtI5vsdX-yUme_sI4pDPDhWqCX5iV-t49YQ21giYk-nYRRo.ilQ9TmQphYEBzEfZ22alAb7CLszjmJ2X8NaF3KW03Bs&dib_tag=se&keywords=4%22%2Bspeakers%2Bheadphone%2Bjack&qid=1767410690&sprefix=4%2Bspeakers%2Bheadphone%2B%2Caps%2C363&sr=8-3&th=1)


[ProTechTrader 4 inch 8 ohm Loud Speaker - 10 w Max / 5 Watt Nominal Square Screw Mount Round Cone for Pinball Arcade Machines M.A.M.E. Video Games DIY Electronics Radios Computers](https://www.amazon.com/ProTechTrader-inch-100mm-Loud-Speaker/dp/B0G3D52FYG/ref=sr_1_1?crid=778HY874AUF1&dib=eyJ2IjoiMSJ9.7gNVbJqq24ak6NVSFZNXws_x94piktodNrrGvHIzL3ImkYpHaEFe4iCtuhmxUISLUv8TUJRGptQAYHRAH-qR3wjGGDvGBLoDr37xwr4WjXgiNTe-V0aCJlmxEGk7AN_XgKvvVUKBxAdlZs2aS78gVg1StAV9W-uZDN22Lr24495EwP2qtOppVqdKlPGBJwrS6ioyFKOw2sOyvcBJ7samzjPpfG2WSDAqEs88GIwwrVo._jIs90-LQK5vMf-e5FD3S0-KTMQnD13cF_ZCODn1J_0&dib_tag=se&keywords=4%22%2Bcomputer%2Bscrew%2Bspeakers&qid=1767410861&sprefix=4%2Bcomputer%2Bscrew%2Bspeakers%2Caps%2C209&sr=8-1&th=1)