function sendMailOnNewResponsesViaMailDraft() {
  var sheet = SpreadsheetApp.getActiveSheet();

  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();

  const heads = data.shift();

  var lastRow = sheet.getLastRow();
  var newFormResponse = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues();

  var row = heads.reduce((o, k, i) => (o[k] = newFormResponse[0][i] || '', o), {});
  console.log(row);
  // return;
  var emailAddress = row['Email Address'];

  const out = [];
  const subjectLine = "QR Code";
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
  try {
    const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

    GmailApp.sendEmail(emailAddress, msgObj.subject, msgObj.text, {
      htmlBody: msgObj.html,
      // bcc: 'a.bbc@email.com',
      // cc: 'a.cc@email.com',
      // from: 'an.alias@email.com',
      // name: 'name of the sender',
      // replyTo: 'a.reply@email.com',
      // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
      attachments: emailTemplate.attachments,
      inlineImages: emailTemplate.inlineImages
    });
    // Edits cell to record email sent date
    out.push([new Date()]);
  } catch (e) {
    // modify cell to record error
    out.push([e.message]);
  }
  console.log(out);

}

function getGmailTemplateFromDrafts_(subject_line) {
  try {
    const drafts = GmailApp.getDrafts();
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    const msg = draft.getMessage();

    const allInlineImages = draft.getMessage().getAttachments({ includeInlineImages: true, includeAttachments: false });
    const attachments = draft.getMessage().getAttachments({ includeInlineImages: false });
    const htmlBody = msg.getBody();

    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});

    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];

    const inlineImagesObj = {};

    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

    return {
      message: { subject: subject_line, text: msg.getPlainBody(), html: htmlBody },
      attachments: attachments, inlineImages: inlineImagesObj
    };
  } catch (e) {
    throw new Error("Oops - can't find Gmail draft");
  }
  function subjectFilter_(subject_line) {
    return function (element) {
      if (element.getMessage().getSubject() === subject_line) {
        return element;
      }
    }
  }
}

function fillInTemplateFromObject_(template, data) {
  let template_string = JSON.stringify(template);
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return JSON.parse(template_string);
}

function escapeData_(str) {
  return str.toString()
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
};
