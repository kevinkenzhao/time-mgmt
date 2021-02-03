# Attendance Management with VBScript

This visual basic script launches a restricted Internet Explorer instance prompting for a QR code to be entered. It assumes that (1) the QR code resolves to a valid HTTP URL containing an employee-specific Google Sheet which will be used to clock in/out as well as estimate net pay and (2) the service account timeclock@nychineseschool.org is signed in on Google Chrome. Upon reading the QR code, a Chrome instance (in kiosk mode) is invoked and navigates to the inputted URL. Once loaded, the end user is presented with a Sheet containing their name and large (Apps Script driven) "Clock In" and "Clock Out" buttons. After 30 seconds, the Chrome process is terminated and the end user is once again presented with the QR code landing page unless the cscript process is terminated via the task manager or the keyword "nycs" is entered.

![alt text](https://github.com/kevinkenzhao/New-York-Chinese-School/blob/master/kiosk-capture.PNG?raw=true)
