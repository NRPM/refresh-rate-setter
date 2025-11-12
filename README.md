# refresh-rate-setter
(THIS PROGRAM WAS MADE BY USING CHATGPT AND CLAUDE AS A SHORT PERSONAL VIBE CODING PROJECT)

Feel free to use it though :)

This is a simple program to set the refresh rate of the display of a Windows 11 laptop depending on it's charging status.

When the laptop is plugged in, the refresh rate is set to the maximum supported refresh rate and when it is unplugged, the refresh rate reduces to the minimum supported refresh rate.

# How to run
Extract the zip file and double click on the .exe file to run the program. (The program always starts minimized to systray)

# Starting the program at system StartUp
1. Make a shortcut of the .exe file.
2. Press Win + R to open the Run window.
3. Type "shell:startup" to open the startup folder.
4. Copy the shortcut into this folder.
5. Now you can freely enable or disable it to start whenever the laptop starts up through the task manager!

(THE PYTHON SCRIPT REQUIRES YOU TO INSTALL PYSTRAY. IT CAN BE INSTALLED USING "pip install pystray")
