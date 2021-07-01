function emailReportToSelf(subject, message) {
  let emailQuotaRemaining = MailApp.getRemainingDailyQuota() - 1;

  let template = HtmlService.createTemplateFromFile('email_template');
  template.message = message;
  template.remainingQuota = emailQuotaRemaining;

  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: subject,
    htmlBody: template.evaluate().getContent(),
  });
}
