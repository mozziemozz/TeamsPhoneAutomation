# Install

Navigate to the folder where you've saved the downloaded files. The folder must contain the following files:

- CreateShortcut.ps1
- TeamsClipBoardDialer.ps1
- PhoneIcon1.ico
- PhoneIcon2.ico

## Run the install script (CreateShortcut.ps1)

Run the script with PowerShell to install the files.

### PowerShell

```powershell
.\CreateShortcut.ps1
```

### Windows Explorer

![run with powershell example](Screenshots/run-with-powershell-example-2023-28-23-23-28-57.png)

## Pin to Taskbar

![pin to taskbar](Screenshots/pin-to-taskbar-2023-32-23-23-32-49.png)

# Usage

Copy any phone number to the clipboard and either click the app on the taskbar or press the configured hotkey. (The default hotkey is **CTRL + SHIFT + F8**).

## Phone number in clipboard

If you have copied a valid phone number, Teams will open and you'll be asked if you want to start a call.

![teams start call](Screenshots/teams-start-call-2023-38-23-23-38-53.png)

### Normalization rule behavior

It's generally advised to only copy a number and no other text. However, you don't need to worry about special characters like white spaces, brackets, hyphens, dots etc. Even illegal zeros which follow after an international prefix are removed. (E.g. +41 (0) 44 123 45 67 is changed to +41 44 123 45 67.)

If your clipboard contains multiple lines, e.g. (when you copied an address which includes a phone number) the script will extract the phone number. This only works if there are more digit chaacters than non-digit characters present.

If your clipboard contains multiple numbers, you'll get a notification and need to copy a single number again.

If the copied number is longer than 16 digits, the number will be disregarded and you'll see a notification.

Copied numbers can be as short as 1 digit. This is because you can also dial internal extensions or emergency numbers.

Numbers starting with 0 (national numbers) will not be converted to E.164 because there's no way to know which country a number is from.

A leading + will be added to numbers which don't already have one and are longer than 4 digits. (E.g. 41441234567 is changed to +41441234567)

Numbers which start with 00, will be converted to E.164 (00 is replaced by +).

## No phone number in clipboard

If you didn't copy a phone number and have something else in the clipboard instead, you'll see an error message and your actual clipboard content.

![balloon tip text](Screenshots/balloon-tip-text-2023-32-24-23-32-11.png)

![balloon tip empty clipboad](Screenshots/balloon-tip-empty-clipboard-2023-31-24-23-31-44.png)

![balloon tip multiple lines](Screenshots/balloon-tip-multiple-lines-2023-12-25-23-12-53.png)

# Buy me a coffee

If you like the work I've done on this app for the Teams community, please consider supporting me via https://github.com/sponsors/mozziemozz.