# New-York-Chinese-School

This visual basic script invokes a restricted Internet Explorer instance prompting for a QR code to be entered. It assumes that the QR code resolves to a valid HTTP URL to a specified Google Sheet which will be used to clock in/out as well as estimate net pay. Upon reading the QR code, a Google Chrome instance in kiosk mode is invoked and will navigate to the inputted URL. Once loaded, the end user is presented with a Sheet containing their name and large (Apps Script enhanced) "Clock In" and "Clock Out" buttons. After 30 seconds, the Chrome process is terminated; the IE instance is uninterrupted unless the cscript process is forcefully terminated via task manager or exit keyword "nycs" is entered.
