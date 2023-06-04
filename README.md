# Wireless-Network-Password-Retrieval-Tool
Wireless Network Password Retrieval Tool: Extracting and Saving Wi-Fi Passwords to Excel

This repository contains a project that serves as a wireless network password retrieval tool developed using Python. The tool utilizes the subprocess module and the openpyxl library to extract and save Wi-Fi passwords to an Excel file.

The code retrieves the list of available Wi-Fi profiles using the netsh wlan show profiles command and collects the profile names. It then proceeds to retrieve the Wi-Fi passwords for each profile using the netsh wlan show profile [profile_name] key=clear command. The collected passwords are stored alongside their respective profile names in an Excel file.

The project provides a convenient way to gather Wi-Fi passwords, making it useful for troubleshooting, network management, or backup purposes. It demonstrates the usage of subprocess calls, string parsing, and Excel file manipulation in Python.
