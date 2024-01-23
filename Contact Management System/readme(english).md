# Contact Management System

## System Introduction

The Contact Management System is a tool designed for individual and small business users, aimed at providing a simple and effective way to manage and maintain contact information. This system allows users to add, search, modify, delete, and filter contacts, as well as sort and analyze contact information.

## Installation Instructions

To run this program, the following libraries are required:

- `PyQt5`
- `openpyxl`
- `matplotlib`
- `smtplib` (usually part of Python's standard library)
- `bypy` (optional, for Baidu Cloud synchronization)
- `playsound`

Example installation command (using pip):

```bash
pip install PyQt5 openpyxl matplotlib playsound
```

## User Guide

To run the program, please follow these steps:

1. Open the command line interface.
2. Navigate to the directory containing the program files.
3. Run the command `python ui.py` to start the user interface.
4. Alternatively, directly execute `ui.py` to launch the system interface.

## Feature Overview

- **Add Contacts**: Allows users to enter detailed information about a contact and save it to the system.
- **Search Contacts**: Users can search for contacts based on different criteria.
- **Modify Contacts**: Allows users to update the information of existing contacts.
- **Delete Contacts**: Users can delete one or all contacts.
- **Sort Contacts**: Sort contacts based on different criteria (such as age, income).
- **Filter Contacts**: Allows users to filter contacts based on specific criteria (such as height, weight).
- **Automatic Birthday Email Sending**: The system automatically detects contacts with upcoming birthdays and sends customized birthday greetings.
- **Visualization Analysis**: Provides a chart to show the proportional distribution of different types of contacts.

## Program Entry

- The main entry file for the program is `ui.py`.
- After launching the program, the user interface will guide users to access all features.

## System Requirements

- Python 3.x
- Compatible with Windows, macOS, and Linux operating systems.

## Additional Information

### Baidu Cloud Binding

Before starting, ensure that Python and pip are successfully installed on your computer.

1. Install the `bypy` library by entering the following command in the command line interface (cmd):

```bash
pip install bypy
```

2. Next, authorize `bypy` to access your Baidu Cloud. Enter the following command in the command line interface:

```bash
bypy list
```

3. After executing this command, a link will appear in the command line interface. Click on this link or copy and paste it into your browser.

4. In the browser, you will be asked to log in to Baidu Cloud and authorize the application. After completing the authorization, you will receive an authorization code.

5. Copy and paste this code back into the command line interface and press Enter. This completes the binding of `bypy`.

You can create a `bypy` folder in "My Cloud>My App Data" on Baidu Cloud. The program will use this folder by default to store and synchronize files.

### Email Binding (Setting Up SMTP Service and Obtaining Authorization Code)

#### QQ Email

##### What is an Authorization Code?

The authorization code is a special password used by QQ Email to log in to third-party clients, applicable for POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV services.

#### Steps to Obtain an Authorization Code

1. Log in to the web version of QQ Email and go to **Settings** -> **Account**.
2. In the "Account" settings, find the **Service Settings** item.
3. Enable **POP3/SMTP service** and verify the security.
4. Send the specified content to the specified number using your phone, and then obtain the authorization code.
5. Enter the 16-digit authorization code in the password box of the third-party client for verification.

#### Important Notes

- Changing the QQ password or independent password will cause the authorization code to expire, requiring reacquisition.


- The authorization code typically does not expire unless the above situation occurs. Once obtained, it can be saved for future use.
- When logging in, use the QQ email address, for example: `12345678@qq.com`.

### NetEase Email

#### Steps to Obtain an Authorization Code

1. Open the web version of NetEase Email and log in.
2. Click **Settings** and choose **POP/SMTP/IMAP options**.
3. In the client protocol interface, choose to enable the relevant protocol (such as POP3/SMTP) and click to activate.
4. In the pop-up window, choose the SMS sending method.
5. Once the system detects successful SMS sending, it will display the client authorization code (a unique 16-letter combination).

#### Important Notes

- NetEase Email allows a maximum of 5 authorization codes at the same time.
- The authorization code is displayed only once when activated, but it can be used for multiple clients.
- It is recommended to generate a new authorization code when needed.

---

**When setting up your own email and authorization code, modify the password and authorization code in the function `send_birthday_email` within the `find_send_email.py` file.**

This is the README file I've prepared. Please use it as a guide for your installation and usage of the Contact Management System.