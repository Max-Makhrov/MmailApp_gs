# MmailApp_gs
Library for sending emails with Google Apps Script. Use `MailApp`, `GmailApp` or `GmailAPI`. The library simpifies the usage of methods.

## Install
Library code is:

`1XHgOqNIXoRxmM7StfIoxxqTVNoEVNpO46UsSSZdXZI4fezxCRIpSYHtB` 🔗[Source](https://script.google.com/u/0/home/projects/1XHgOqNIXoRxmM7StfIoxxqTVNoEVNpO46UsSSZdXZI4fezxCRIpSYHtB/edit)

If you would like to use `Gmail API`, please enable it as [advanced Google Service](https://developers.google.com/apps-script/guides/services/advanced).

## Try

Please see how it works in a visual way:

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
    htmlBody: '<b>💪Bold Hello</b><br>How are you?',
    noreply: 1,
    replyto: 'test@test.test',
    // replytoname: 'Maaan', // works for Gmail API
    // attachments: array,
    // inlineimages: object
  }
  try {
    // try using MailApp, or GmailApp
    var response = MmailApp.send(options);
    console.log(response);
  } catch(err) {
    console.log(err);
  }
}
```

If you want to force the use of Gmail API:

 ```
  // this will forse the use of Gmail API
  MMailApp.sendGmailApi(options);
 ```
 
 ## Author
@max__makhrov

![CoolTables](https://raw.githubusercontent.com/cooltables/pics/main/logos/ct_logo_small.png) [cooltables.online](https://www.cooltables.online/)
