# MmailApp_gs
Library for sending emails with Google Apps Script

## Install
Library code is:

`1XHgOqNIXoRxmM7StfIoxxqTVNoEVNpO46UsSSZdXZI4fezxCRIpSYHtB` 🔗[Source](https://script.google.com/u/0/home/projects/1XHgOqNIXoRxmM7StfIoxxqTVNoEVNpO46UsSSZdXZI4fezxCRIpSYHtB/edit)

## Try

Please see how it works in visual way:

1. Copy [⚡ Automation Samples](https://docs.google.com/spreadsheets/d/1HwUaZk86BtrPdQ1RYILwTcRwJUUClgqtAPEpMAsX0y8/edit#gid=1396833470) project file. Go to menu `File > Make a Copy`.
2. Change settings on the sheet: `_sendmail_`. 
3. Use the menu `💌 Mail > 📮 Send Emails` to test the script.
4. Success: you'll send emails and see the logs in the last rows of sheet: `📜 logs`.

## Use

Please use this sample code with all possible settings:

```
function test_getMailApp() {
  var options = {
    to: 'makhrov.max+spam@gmail.com', // required
    cc: 'max0637859167@gmail.com',
    bcc: 'max0637859167@gmail.com,makhrov.max+spam@gmail.com',
    name: '🎩Mad Hatter',
    // from: 'max0637859167@gmail.com', // for paid accounts + alias
    subject: '🥸Test Test Test',
    body: '💪 body!',
    htmlBody: '<b>💪Bold Hello</b><br>How are you?'
    noreply: 1,
    replyto: 'test@test.test',
    // replytoname: 'Maaan', // works for Gmail API
    // attachments: array,
    // inlineimages: object
  }
  var response = MmailApp.send(options);
}
```
