var MMailApp = this;

/** send
 * ### ⚙️Options 
 * 
 * Semple options:
 * 
 * `{to: `email`, subject: "hi", htmlbody: "<b>Bold!</b>"}`
 * 
| **Parameter**   | **Type** | **Notes**                                                             |
|:----------------|:---------|:----------------------------------------------------------------------|
| `to`            | String   | comma-separated, also may use: `recipient`, `recipients`              |
| `cc`            | String   | comma-separated                                                       |
| `bcc`           | String   | comma-separated                                                       |
| `subject`       | String   |                                                                       |
| `body`          | String   |                                                                       |
| `htmlbody`      | String   |                                                                       |
| `name`          | String   | default: the user's name                                              |
| `from`          | String   | for paid accounts. One of `var aliases = GmailApp.getAliases();`      |
| `replyto`       | String   |                                                                       |
| `noreply`       | Integer  | 1 =`true`, for paid accounts                                          |
| `inlineimages`  | Object   | `{name: blob}`, use in HtmlBody: `inline <img src='cid:name'> image!` |
| `attachments`   | Array    | `[blob, blob]`                                                        |
 * 
 * ### Author
 * `@max__makhrov` 
 * 
 * ![CoolTables](https://raw.githubusercontent.com/cooltables/pics/main/logos/ct_logo_small.png) 
 *  [cooltables.online](https://www.cooltables.online/)
 * 
 */
function send(options) {
  /** 
   * 
   * 1. Try MailApp
   * 2. Try GmailApp
   * 3. Try Gmail API
   * 
   * if `from` parameter is set, the script will try GmailApp
   * 
   */
  var settings = {
    min_quota: 2
  }

  // let use of different naming conventions
  setOptionsNamings_(options);

  // get remaining quota
  var quota = getRemainingMailQuota_();
  if (quota <= settings.min_quota) {
    console.log('Regular sendMail Quota is exceeded. Use Gmail API');
    if (typeof Gmail == 'undefined') {
      throw '🧐you are out of your daily Quota. 👌Fix: Please enable Gmail API Advanced Service.'
    }
    var res = sendGmailApi(options);
    return res;
  }

  if (options.from && options.from !== '') {
    console.log('use GmailApp');
    sendGmail_(options);
  } else {
    console.log('use MailApp');
    sendMail_(options);
  }


  /** returns quota in case GmialApp of MailApp was used */
  return {
    ramaining_quota: getRemainingMailQuota_()
  }
}

/**
 * let user use 
 * 
 *   `options.htmlbody` OR `options.htmlBody`
 * 
 * and other naming conventions
 */
function setOptionsNamings_(options) {
  options.to = options.to || options.recipient || options.recipients;
  options.htmlbody = options.htmlbody || options.htmlBody || options.html_body;
  options.replyto = options.replyto || options.replyTo || options.reply_to;
  options.replytoname = options.replytoname || options.replyToName || options.reply_to_name;
  options.noreply = options.noreply || options.noReply || options.no_reply;
  options.inlineimages = options.inlineimages || options.inlineImages || options.inline_images;
}


/**
 * Get and log your daily mail quota
 */
function getRemainingMailQuota_() {
  // https://stackoverflow.com/questions/2744520
  var quota = MailApp.getRemainingDailyQuota();
  // console.log('You can send ~ ' + quota  + ' emails this day.');
  return quota;
}



//   __  __       _ _                      
//  |  \/  |     (_) |   /\                
//  | \  / | __ _ _| |  /  \   _ __  _ __  
//  | |\/| |/ _` | | | / /\ \ | '_ \| '_ \ 
//  | |  | | (_| | | |/ ____ \| |_) | |_) |
//  |_|  |_|\__,_|_|_/_/    \_\ .__/| .__/ 
//                            | |   | |    
//                            |_|   |_|    

// function test_getMailApp() {
//   var test_blobs = getTestBlobs_();
//   var options = {
//     to: 'makhrov.max@gmail.com', // required
//     cc: 'max0637859167@gmail.com',
//     bcc: 'max0637859167@gmail.com,makhrov.max@gmail.com',
//     name: '🎩Mad Hatter',
//     from: 'max0637859167@gmail.com', // works only if for paid accounts
//     subject: '🥸Test Test Test',
//     body: '💪 body!',
//     htmlBody: test_blobs.htmlBody, // '<b>💪Bold Hello</b><br>How are you?'
//     noreply: 1,
//     replyto: 'test@test.test',
//     attachments: test_blobs.array,
//     inlineimages: test_blobs.object
//   }
//   var response = sendMail_(options);
// }

/** sendMail_
 *  
 * As descubed here:
 *  https://developers.google.com/apps-script/reference/mail/mail-app#sendemailrecipient,-subject,-body,-options
 */
function sendMail_(options) {
  // Warning
  // noReply is only for paid accounts
  //
  // Note
  // You cannot use `from` with this method

  // let use of different naming conventions
  setOptionsNamings_(options);

  /** body */
  if (!options.body || options.body === '') {
    options.body = '<this message was sent automatically without email body>';
  }

  /** subject */
  if (!options.subject || options.subject === '') {
    options.subject = '(No Subject)'
  }

  /** other options */
  var addProperty_ = function(main, donor, mainKey, donorKey) {
    var val = donor[donorKey];
    if (donorKey === 'noreply' ) {
      if (val) {
        main[mainKey] = true;
        return;
      }
    }
    if ( val && val !== '' ) {
      main[mainKey] = val;
    }  
  }
  var mail_options = { };
  var keys = {
    attachments: 'attachments',
    bcc: 'bcc',
    cc: 'bcc',
    htmlbody: 'htmlBody',
    inlineimages: 'inlineImages',
    name: 'name',
    noreply: 'noReply',
    replyto: 'replyTo'
  }
  for (var k in keys) {
    addProperty_(
      mail_options,
      options,
      keys[k],
      k );
  }

  // console.log(JSON.stringify(mail_options, null, 4))

  MailApp.sendEmail(
    options.to, 
    options.subject, 
    options.body,
    mail_options
  );

  return 0;
}



//    _____                 _ _                      
//   / ____|               (_) |   /\                
//  | |  __ _ __ ___   __ _ _| |  /  \   _ __  _ __  
//  | | |_ | '_ ` _ \ / _` | | | / /\ \ | '_ \| '_ \ 
//  | |__| | | | | | | (_| | | |/ ____ \| |_) | |_) |
//   \_____|_| |_| |_|\__,_|_|_/_/    \_\ .__/| .__/ 
//                                      | |   | |    
//                                      |_|   |_|    

// function test_getGmailApp() {
//   var test_blobs = getTestBlobs_();
//   var options = {
//     to: 'makhrov.max@gmail.com', // required
//     cc: 'max0637859167@gmail.com',
//     bcc: 'max0637859167@gmail.com,makhrov.max@gmail.com',
//     name: '🎩Mad Hatter',
//     from: 'max0637859167@gmail.com', // works only if for paid accounts
//     subject: '🥸Test Test Test',
//     body: '💪 body!',
//     // htmlBody: test_blobs.htmlBody, // '<b>💪Bold Hello</b><br>How are you?'
//     noreply: 1,
//     replyto: 'test@test.test',
//     attachments: test_blobs.array,
//     inlineimages: test_blobs.object
//   }
//   var response = sendGmail_(options);
// }

/** sendGmail_
 *  
 * As descubed here:
 *  https://developers.google.com/apps-script/reference/gmail/gmail-app#sendemailrecipient,-subject,-body,-options
 */
function sendGmail_(options) {
  // Warning
  // noReply is only for paid accounts
  //
  // Warning!
  // Emojis does not work with this method
  //
  // Note
  // Usage of `from` is limited
  // You need to create an alias first
  // https://support.google.com/mail/answer/22370?hl=en
  // you must have a paid account
  // and create alias.
  // To test your aliases:
  // var aliases = GmailApp.getAliases();

  // let use of different naming conventions
  setOptionsNamings_(options);

  /** body */
  if (!options.body || options.body === '') {
    options.body = '<this message was sent automatically without email body>';
  }

  /** subject */
  if (!options.subject || options.subject === '') {
    options.subject = '(No Subject)'
  }

  /** other options */
  var addProperty_ = function(main, donor, mainKey, donorKey) {
    var val = donor[donorKey];
    if (donorKey === 'noreply' ) {
      if (val) {
        main[mainKey] = true;
        return;
      }
    }
    if ( val && val !== '' ) {
      main[mainKey] = val;
    }  
  }
  var mail_options = { };
  var keys = {
    attachments: 'attachments',
    bcc: 'bcc',
    cc: 'bcc',
    'from': 'from',
    htmlbody: 'htmlBody',
    inlineimages: 'inlineImages',
    name: 'name',
    noreply: 'noReply',
    replyto: 'replyTo'
  }
  for (var k in keys) {
    addProperty_(
      mail_options,
      options,
      keys[k],
      k );
  }

  // console.log(JSON.stringify(mail_options, null, 4))

  GmailApp.sendEmail(
    options.to, 
    options.subject, 
    options.body,
    mail_options);

  return 0;
}




//    _____                 _ _            _____ _____ 
//   / ____|               (_) |     /\   |  __ \_   _|
//  | |  __ _ __ ___   __ _ _| |    /  \  | |__) || |  
//  | | |_ | '_ ` _ \ / _` | | |   / /\ \ |  ___/ | |  
//  | |__| | | | | | | (_| | | |  / ____ \| |    _| |_ 
//   \_____|_| |_| |_|\__,_|_|_| /_/    \_\_|   |_____|
// function test_getGmailApiSets() {
//   var test_blobs = getTestBlobs_();
//   var options = {
//     to: 'makhrov.max@gmail.com', // required
//     cc: 'max0637859167@gmail.com',
//     bcc: 'max0637859167@gmail.com,makhrov.max@gmail.com',
//     name: '🎩Mad Hatter',
//     // from: 'max0637859167@gmail.com', // works only if for paid accounts
//     subject: '🥸Test Test Test',
//     body: '💪 body!',
//     htmlBody: test_blobs.htmlBody, // '<b>💪Bold Hello</b><br>How are you?'
//     noReply: 1,
//     replyTo: 'test@test.test',
//     replyToName: '🦸Stranger',
//     attachments: test_blobs.array,
//     inlineImages: test_blobs.object
//   }
//   var response = sendGmailApi_(options)
//   console.log(JSON.stringify(response))
// }


/** sendGmailApi
 * ### Warning
 * ⚠️Enable Gmail 
 * [advanced service](https://developers.google.com/apps-script/guides/services/advanced)
 *  first in order to use this code
 * 
 * ### ⚙️Options 
 * 
 * Semple options:
 * 
 * `{to: email, subject: "hi", htmlbody: "<b>Bold!</b>"}`
 * 
| **Parameter**   | **Type** | **Notes**                                                             |
|:----------------|:---------|:----------------------------------------------------------------------|
| `to`            | String   | comma-separated, also may use: `recipient`, `recipients`              |
| `cc`            | String   | comma-separated                                                       |
| `bcc`           | String   | comma-separated                                                       |
| `subject`       | String   |                                                                       |
| `body`          | String   |                                                                       |
| `htmlbody`      | String   |                                                                       |
| `name`          | String   | default: the user's name                                              |
| `from`          | String   | for paid accounts. One of `var aliases = GmailApp.getAliases();`      |
| `replyto`       | String   |                                                                       |
| `noreply`       | Integer  | 1 = no reply = true                                                   |
| `inlineimages`  | Object   | `{name: blob}`, use in HtmlBody: `inline <img src='cid:name'> image!` |
| `attachments`   | Array    | `[blob, blob]`                                                        |
🕵There's also a hidden option: `replyToName`. This option is for Gmail API only
 * 
 * 
 * ### Return Sample
 ```
{"labelIds":["UNREAD","SENT","INBOX"],
  "id":"1831c4c0aa922712",
  "threadId":"1831c4c0aa922712",
  "url":"https://mail.google.com/mail/u/0/#inbox/1831c4c0aa922712"}
 ``` 
 * 
 * ### Author
 * `@max__makhrov` 
 * 
 * ![CoolTables](https://raw.githubusercontent.com/cooltables/pics/main/logos/ct_logo_small.png) 
 *  [cooltables.online](https://www.cooltables.online/)
 * 
 */
function sendGmailApi(options) {
  // Note
  // Usage of `from` is limited
  // You need to create an alias first
  // https://support.google.com/mail/answer/22370?hl=en
  // you must have a paid account
  // and create alias.
  // To test your aliases:
  // var aliases = GmailApp.getAliases();
  var rfcString = getGmailApiString_(options);
  // console.log(rfcString);
  var result = sendGmailApiByRaw_(rfcString);
  return result;
}


/**
 * send Gmail API email
 * 
 * @param {string} raw - RFC 2822 formatted and base64url encoded email
 * 
 */
function sendGmailApiByRaw_(raw) {
  var raw_encoded = Utilities.base64EncodeWebSafe(raw);
  // console.log(raw_encoded);
  var response = Gmail.Users.Messages.send({raw: raw_encoded}, "me");
  if (response.threadId) {
    response.url = 'https://mail.google.com/mail/u/0/#inbox/' + response.threadId;
  }
  return response;
}


/**
 * get raw RFC 2822 formatted and base64url encoded email
 * 
 * ## Credits
 * [😺github/muratgozel/MIMEText](https://github.com/muratgozel/MIMEText) — RFC 2822 compliant raw email message generator
 * 
 * [♨️Emoji in email subject with Apps Script?](https://stackoverflow.com/a/66088350/5372400) - sample script by Tanaike
 */
function getGmailApiString_(options) {
  // let use of different naming conventions
  setOptionsNamings_(options);

  /** parameters */
  var boundary = 'happyholidays';
  var boundary2 = 'letthemerrybellskeepringing';
  var dd = '--';
  var header = ['MIME-Version: 1.0'];
  var body = [];

  /** for emojis to go */
  var getEncoded_ = function(str) {
    return Utilities.base64Encode(
        str,
        Utilities.Charset.UTF_8);
  }
  /** for emojis to go in names */
  var getEncodedString_ = function(str) {
    return '=?UTF-8?B?' + getEncoded_(str) + '?=';
  }
  /** add <tags> */
  var addTags_ = function(str) {
    return '<' + str + '>';
  }
  /** for emails */
  var getParsedEmails_ = function(str) {
    var arr = str.split(',').map(addTags_);
    return arr.join(', ');
  }

  /** 1️⃣From */
  var fromemail = Session.getEffectiveUser().getEmail();
  if (options.from) {
    fromemail = options.from;
  }
  var i_from = addTags_(fromemail);
  if (options.name && options.name !== '') {
    i_from = getEncodedString_(options.name) + 
      ' ' + i_from;
  }
  header.push('From: ' + i_from);

  /** 2️⃣Reply To & No Reply */
  if (options.replyto && options.replyto !== '') {
    var replyto = addTags_(options.replyto);
    if (options.replytoname && options.replytoname !== '') {
      replyto = getEncodedString_(options.replytoname) + 
        ' ' + replyto;
    }
    header.push('Reply-To: ' + replyto);
  } else if (options.noreply) {
    header.push('Reply-To: No Reply <noreply@test.test>');
  }

  /** 3️⃣ To */
  if (!options.to || options.to === '') {
    throw "😮resepient is missing. Please add like this: `{to: 'mail@test.test'}`"
  }
  header.push('To: ' + getParsedEmails_(options.to));

  /** 4️⃣ Cc & Bcc */
  if (options.cc && options.cc !== '') {
    header.push('Cc: ' + getParsedEmails_(options.cc));
  }
  if (options.bcc && options.bcc !== '') {
    header.push('Bcc: ' + getParsedEmails_(options.bcc));
  }

  /** 5️⃣ Subject */
  if (!options.subject || options.subject === '') {
    options.subject = '(No Subject)'
  }
  header.push('Subject: ' + getEncodedString_(options.subject));
  

  /** 6️⃣ Other Headers & delimiters */
  body.push('Content-Type: multipart/mixed; boundary=' + boundary);
  body.push('');
  body.push(dd + boundary);
  body.push('Content-Type: multipart/alternative; boundary=' + boundary2);
  body.push('');

  /** 7️⃣ body & htmlBody */
  body.push(dd + boundary2);
  body.push('Content-Type: text/plain; charset=UTF-8');
  body.push('Content-Transfer-Encoding: base64');
  body.push('');
  if (!options.body || options.body === '') {
    options.body = '<this message was sent automatically without email body>';
  }
  body.push(getEncoded_(options.body));
  if (options.htmlbody && options.htmlbody !== '') {
    body.push(dd  +boundary2);
    body.push('Content-Type: text/html; charset=UTF-8');
    body.push('Content-Transfer-Encoding: base64');
    body.push('');
    body.push(getEncoded_(options.htmlbody));
  }

  /** 8️⃣ Attachments */
  if (options.attachments && options.attachments !== '') {
    body.push(dd + boundary2 + dd);  
    var attach;
    for (var i = 0; i < options.attachments.length; i++) {
      attach = options.attachments[i];
      body.push(dd + boundary);
      body.push('Content-Type: ' + attach.getContentType() + '; charset=UTF-8');
      body.push('Content-Transfer-Encoding: base64');
      body.push('Content-Disposition: attachment;filename="' + attach.getName() + '"');
      body.push('');
      body.push(Utilities.base64Encode(attach.getBytes()));
    }
  }

  /** 9️⃣ Inline Images */
  if (options.inlineimages && options.inlineimages !== '') {
    body.push(dd + boundary2 + dd);  
    var inline;
    for (var k in options.inlineimages) {
      inline = options.inlineimages[k];
      body.push(dd + boundary);
      body.push('Content-Type: ' + inline.getContentType() + '; name=' + inline.getName());
      body.push('Content-Transfer-Encoding: base64');
      body.push('X-Attachment-Id: ' + k);
      body.push('Content-ID: <' + k + '>');
      body.push('');
      body.push(Utilities.base64Encode(inline.getBytes()));
    }
  }

  /** 🔟 Result */
  body.push('');
  body.push(dd  + boundary + dd);
  var result = header.concat(body).join('\r\n');
  return result;
}                                                 




//   _______        _       
//  |__   __|      | |      
//     | | ___  ___| |_ ___ 
//     | |/ _ \/ __| __/ __|
//     | |  __/\__ \ |_\__ \
//     |_|\___||___/\__|___/
/**
 * this function demonstrates 
 * how to get blobs by URL
 * for your email
*/
function getTestBlobs_() {
  var source = 'https://raw.githubusercontent.com/';
  var options = {
    urls: {
      'coolTablesLogo': 
        source + 'cooltables/pics/main/logos/ct_logo_small.png',
      'partyFace':
        source + 'Max-Makhrov/myFiles/master/partyface.png'
    },
    html: "🔍 inline CoolTables Logo<img src='cid:coolTablesLogo'> image! <br>" +
      "Hoooorrraaay! <img src='cid:partyFace'>"
  }

  var blobsArray = [], blobsObject = {};
  var blob;
  for (var k in options.urls) {
    blob = UrlFetchApp
      .fetch(options.urls[k])
      .getBlob()
      .setName(k);
    blobsArray.push(blob);
    blobsObject[k] = blob;
  }

  return {
    array: blobsArray, // for attachments
    object: blobsObject, // for inline images
    htmlBody: options.html
  };
}
