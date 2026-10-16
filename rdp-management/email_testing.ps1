$SCRIPT:SMGUserAccessType = Read-Host "Enter TS or HTML5"



$html = "<p class=MsoNormal>The Sewell SSO Remote Access system uses the Okta Verify app to provide two-factor authentication.&nbsp; This system is easy to use while maintaining a high level of security.</p>"

$html += "<p class=MsoNormalCxSpMiddle style='margin-top:0in;margin-right:0in;margin-bottom:8.0pt;margin-left:.5in;mso-add-space:auto;text-indent:-.25in;line-height:105%;mso-list:l0 level1 lfo2'><![if !supportLists]><span style='mso-fareast-font-family:Calibri'><span style='mso-list:Ignore'>1)<span style='font:7.0pt `"Times New Roman`"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span></span><![endif]>Download the Okta Verify app on your iOS or Android device.</p>"

If ($SCRIPT:SMGUserAccessType -eq "TS") {
    $html += "<p class=MsoNormalCxSpMiddle style='margin-top:0in;margin-right:0in;margin-bottom:8.0pt;margin-left:.5in;mso-add-space:auto;text-indent:-.25in;line-height:105%;mso-list:l0 level1 lfo2'><![if !supportLists]><span style='mso-fareast-font-family:Calibri'><span style='mso-list:Ignore'>2)<span style='font:7.0pt `"Times New Roman`"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span></span><![endif]>On a Windows 10 computer, open Internet Explorer (not Microsoft Edge) or Google Chrome and login to <a href=`"https://sewell.okta.com`">https://sewell.okta.com</a> using your Sewell Windows credentials.</p>"

    $EmailSubject = "Sewell SSO Remote Access - Windows"
} ElseIf ($SCRIPT:SMGUserAccessType -eq "HTML5") {
    $html += "<p class=MsoNormalCxSpMiddle style='margin-top:0in;margin-right:0in;margin-bottom:8.0pt;margin-left:.5in;mso-add-space:auto;text-indent:-.25in;line-height:105%;mso-list:l0 level1 lfo2'><![if !supportLists]><span style='mso-fareast-font-family:Calibri'><span style='mso-list:Ignore'>2)<span style='font:7.0pt `"Times New Roman`"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span></span><![endif]>On a Mac computer, open Safari and login to <a href=`"https://sewell.okta.com`">https://sewell.okta.com</a> using your Sewell Windows credentials.</p>"

    $EmailSubject = "Sewell SSO Remote Access - Mac"
}

$html += "<p class=MsoNormalCxSpMiddle style='margin-top:0in;margin-right:0in;margin-bottom:8.0pt;margin-left:.5in;mso-add-space:auto;text-indent:-.25in;line-height:105%;mso-list:l0 level1 lfo2'><![if !supportLists]><span style='mso-fareast-font-family:Calibri'><span style='mso-list:Ignore'>3)<span style='font:7.0pt `"Times New Roman`"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span></span><![endif]>Locate the Pulse Secure Remote Access button.</p>"

$html += "<p class=MsoNormalCxSpMiddle style='margin-top:0in;margin-right:0in;margin-bottom:8.0pt;margin-left:.5in;mso-add-space:auto;text-indent:-.25in;line-height:105%;mso-list:l0 level1 lfo2'><![if !supportLists]><span style='mso-fareast-font-family:Calibri'><span style='mso-list:Ignore'>4)<span style='font:7.0pt `"Times New Roman`"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span></span><![endif]>Click on the button to go to the Okta Verify portion of the login.</p>"

$html += "<p class=MsoNormalCxSpMiddle style='margin-top:0in;margin-right:0in;margin-bottom:8.0pt;margin-left:.5in;mso-add-space:auto;text-indent:-.25in;line-height:105%;mso-list:l0 level1 lfo2'><![if !supportLists]><span style='mso-fareast-font-family:Calibri'><span style='mso-list:Ignore'>5)<span style='font:7.0pt `"Times New Roman`"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span></span><![endif]>Open the Okta Verify app on your iOS or Android device.&nbsp; It will instruct you to take a picture of the QR code on the screen.&nbsp; This will enroll your device and enable it to be used for secondary authentication to the Sewell Remote Access site.</p>"

$html += "<p class=MsoNormalCxSpMiddle style='margin-top:0in;margin-right:0in;margin-bottom:8.0pt;margin-left:.5in;mso-add-space:auto;text-indent:-.25in;line-height:105%;mso-list:l0 level1 lfo2'><![if !supportLists]><span style='mso-fareast-font-family:Calibri'><span style='mso-list:Ignore'>6)<span style='font:7.0pt `"Times New Roman`"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span></span><![endif]>Once enrolled, the Okta Verify app can be used two ways to verify your identity.&nbsp; The next time you click on the Pulse Secure Remote Access button at <a href=`"https://sewell.okta.com`">https://sewell.okta.com</a>, you will be presented with two buttons: $([char]34)Send Push$([char]34) and $([char]34)Or enter code$([char]34).&nbsp; $([char]34)Send Push$([char]34) will send a notification to your phone that you can either approve or deny that you are logging in.&nbsp; $([char]34)Or enter code$([char]34) opens a field where you will need input a password token that you get by opening the Okta Verify app and retrieving the code.</p>"

$html += "<p class=MsoNormalCxSpMiddle style='margin-top:0in;margin-right:0in;margin-bottom:8.0pt;margin-left:.5in;mso-add-space:auto;text-indent:-.25in;line-height:105%;mso-list:l0 level1 lfo2'><![if !supportLists]><span style='mso-fareast-font-family:Calibri'><span style='mso-list:Ignore'>7)<span style='font:7.0pt `"Times New Roman`"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></span></span><![endif]>Once you verify you are the correct person logging in, you will be logged on to the Sewell Remote Access site.</p>"

$html += "<br /><p class=MsoNormal>Please reply to this email if you have any issues or questions relating to remote access. Thank you and have a lovely day.</p>"

$html += "<p class=MsoNormal>Sewell IT Infrastructure</p>"


# Use the HTML table and these additional variables to send the results as an email #
$SMTPServer = "mail2.sewell.com"
$EmailBody = $html 
Send-MailMessage -SmtpServer $SMTPServer -From "remote.access.support@sewell.com" -To "shamilton@sewell.com" -Subject $EmailSubject -Body $EmailBody -BodyAsHtml