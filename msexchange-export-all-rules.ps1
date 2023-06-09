# Script to export all rules of all mailboxes in Microsoft Exchange

# $logins = Get-Mailbox -Resultsize Unlimited | Select PrimarySmtpAddress
# $login = $logins.PrimarySmtpAddress

$login = gc 'D:\allusers.txt'
$result = foreach ($l in $login) {

    Get-InboxRule -mailbox $l | SELECT @{name="login";expression={$l}}, name, enabled, @{name="from";expression={$_.from[0].ToString()}}, @{name="SentTo";expression={$_.SentTo[0].ToString()}}, @{name="MoveToFolder";expression={$_.MoveToFolder[0].ToString()}}, Description, @{name="SubjectContainsWords";expression={$_.SubjectContainsWords[0].ToString()}}, @{name="ExceptIfSubjectContainsWords";expression={$_.ExceptIfSubjectContainsWords[0].ToString()}}, @{name="BodyContainsWords";expression={$_.BodyContainsWords[0].ToString()}}, @{name="ExceptIfBodyContainsWords";expression={$_.ExceptIfBodyContainsWords[0].ToString()}}, @{name="SubjectOrBodyContainsWords";expression={$_.SubjectOrBodyContainsWords[0].ToString()}}, @{name="ExceptIfSubjectOrBodyContainsWords";expression={$_.ExceptIfSubjectOrBodyContainsWords[0].ToString()}}

}

$result | Export-Csv D:\rulesExchange.csv -Delimiter ';' -Encoding UTF8 -NoTypeInformation
