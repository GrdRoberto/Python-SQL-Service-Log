***LogGenPez Installation and Service Setup***


This document outlines the steps to create an executable from the LogGenPez Python script and set it up as a Windows service using NSSM (Non-Sucking Service Manager).

Prerequisites
Ensure you have Python and PyInstaller installed on your system.
Download NSSM from NSSM's official website.
Step 1: Create Executable
Open your command prompt.

Navigate to the directory containing the LogGenPez.py file.

Run the following command to create a single executable file:


```pyinstaller --onefile --noconsole LogGenPez.py```
Once the execution completes, locate the generated .exe file in the dist folder inside your project directory.

Step 2: Install NSSM
Extract the NSSM archive you downloaded.
Navigate to the win64 folder in the extracted NSSM directory.
Step 3: Install the Executable as a Windows Service
Open the command prompt in the win64 directory by entering this path in Windows Explorer and then typing cmd in the address bar.

Install the service with the following command:

```nssm install "nameservice" "path\to\your\exe\location.exe" "path\to\your\exe\location.exe"```
Replace "path\to\your\exe\location.exe" with the actual path to the created executable.

Follow the prompt to configure service options as necessary and confirm the installation.

Step 4: Managing the Service
To delete the service, use the following command in the command prompt administrator:

```sc delete nameservice```
