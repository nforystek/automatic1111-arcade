# automatic1111-arcade

!\[Stable Diffusion Automatic1111 Arcade](./sd-arcade.jpg)



An app written in VB6 and C++ for Windows 11 that runs as a legacy Kiosk and intends to be a coin operated Text-to-Image Arcade setup using Stable Diffusion Automatic1111 API (on localhost, without internet nor mouse).

Users put in 0.25 cents and get 3 credits.  Each image generated costs 2 credits, and to vote on any one generated image to be in the top 12 is 1 credit.

The following Automatic1111 extensions are invovled/supported:
stable-diffusion-webui-auto-tls-https	https://github.com/papuSpartan/stable-diffusion-webui-auto-tls-https.git
stable-diffusion-webui-nsfw-censor	https://github.com/AUTOMATIC1111/stable-diffusion-webui-nsfw-censor.git

The main application is written in Visual Basic 6 and requires "Microsoft ActiveX Data Objects 2.8 Library" (you can get with MDAC\_TYP.EXE and I'm assuming in a commercial world you'll need a Microsoft Access License as well) and there are two supporing C++ DLL's required, one of which is a Redirect.DLL that allows Visual Basic 6 to open console applications (not visible) and read/write to the console app such as if typing into "cmd.exe," and the second DLL is the 64bit version of inpout32.dll (or inpoutx64.dll) which is used to read IO data logics on a parallel port that this application does to hardware detect a coin box's switch when a coin is put in it.  The two DLL's source code are included in this repository.

IF YOU INTEND TO TRY TO RUN THIS PROGRAM, PLEASE READ THE COMPILE CONDITIONS COMMENTS IN THE SOURCE CODE CAREFULLY AS IT IS POSSIBLE TO LOCK YoU OUT OF THE SYSTEM LEAST THE POWER BUTTON AVAILABLE AND RECOVERY USED.

The application runs as a Windows Shell with all restrictions nessisary to be a public Kiosk with a keyboard.  In a full production envrionment (inside a coin operated arcade box housing) auto logon to a password-less administrator account is the only thing required that does not setup with the code.

