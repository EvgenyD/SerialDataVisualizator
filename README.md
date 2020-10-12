# SerialDataVisualizator
Read and plot numbers from serial port

the app supposed to simplify acquisition and logging of numerica data from serial port.
it requires no installation, no aexternal drivers, dependencies, besides .Net4.0+.
the functionality is kept to minimum so the final size of the exe is small enough to be put directly into the memory of some MCU.
Most of the funcitionality is straightforward.

# 1. Incoming data
It can process any incoming string ending with \n and, if needed, strarting with custom set of chars e.g. ">>"
The incomming message can contain whatever charracters but only numbers are considered and plotted:
each number should be surrounded by these symbols: {" ", ":", ";", ",", "/", "\t", "_", "="} e.g. ">>12;3.5; 3 123123.0,16/255\n"

# 2. Logging
when checked "LOG" the full incoming string is written to the txt file generated on start or when the button "new" is pressed.
filename is generated automaticaly in the /Logs folder next to the exe file.
To open log file you can doubleclicked on the file name in the list.
Right click on the list will open the log folder

# 3. Plot
All the incomming numbers (up to 30) are plotted, you may choose which one to show by checking/unchecking boxes with corresponding color.
Doubleclicking on the plot will clean all the curves and restart plotting from the beginning.
Every 2 seconds, the plots is autorescaled in order to accomodate all the visible plots.

# 4. Outcoming data.
The software allows sending custom command to the serial port in two ways. 
First, the message canbe directly written in the textbox and send using the button "send"
Second, the message can contain predefined number of parameters (numerical, 0-255, or text). this parameters are sombined into the outcoming string using selected separator and saved for later use. each parameter can be skipped for the currend message using corresponding checkbox.
Also, if needed, the message can be sent periodically with period in seconds



