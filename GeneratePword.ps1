function New-Password	{
<#
.SYNOPSIS
  Displays a complex password.

.DESCRIPTION
  Displays a complex password. Output includes password, and phonetic breakdown of the password.

.NOTES
  Version                 : 1.3
  Wish List               :
  Rights Required         : No special rights required
                          : If script is not signed, ExecutionPolicy of RemoteSigned (recommended) or Unrestricted (not recommended)
                          : If script is signed, ExecutionPolicy of AllSigned (recommended), RemoteSigned, 
                            or Unrestricted (not recommended)
  Sched Task Required     : No
  Lync/Skype4B Version    : N/A
  Author/Copyright        : Â© Pat Richard, Skype for Business MVP - All Rights Reserved
  Email/Blog/Twitter      : pat@innervation.com   http://www.ucunleashed.com @patrichard
  Dedicated Post          : http://www.ucunleashed.com/915
  Disclaimer              : You running this script means you won't blame me if this breaks your stuff. This script is
                            provided AS IS without warranty of any kind. I disclaim all implied warranties including, 
                            without limitation, any implied warranties of merchantability or of fitness for a particular
                            purpose. The entire risk arising out of the use or performance of the sample scripts and 
                            documentation remains with you. In no event shall I be liable for any damages whatsoever 
                            (including, without limitation, damages for loss of business profits, business interruption,
                            loss of business information, or other pecuniary loss) arising out of the use of or inability
                            to use the script or documentation. 
  Acknowledgements        : 
  Assumptions             : ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
  Limitations             : 
  Known issues            : None yet, but I'm sure you'll find some!                        

.LINK
  <blockquote class="wp-embedded-content" style="display: none;" data-secret="c1CG7b6OCu"><a href="https://www.ucunleashed.com/915">Function: New-Password – Creating Passwords with PowerShell</a></blockquote><iframe width="500" height="286" title="“Function: New-Password – Creating Passwords with PowerShell” — UC Unleashed" class="wp-embedded-content" src="https://www.ucunleashed.com/915/embed#?secret=c1CG7b6OCu" frameborder="0" marginwidth="0" marginheight="0" scrolling="no" data-secret="c1CG7b6OCu" security="restricted" sandbox="allow-scripts"></iframe>
  
.EXAMPLE
  New-Password -Length <integer>

  Description
  -----------
  Creates a password of the defined length

.EXAMPLE
  New-Password -Length <integer> -ExcludeSymbols

  Description
  -----------
  Creates a password of the defined length, but does not utilize the following characters: !$%^-_:;{}<># &@]~

.INPUTS
  This function does support pipeline input.
#>
  #Requires -Version 3.0
    
	[CmdletBinding(SupportsShouldProcess = $true)]
	param(
		#Defines the length of the desired password
		[Parameter(ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
		[ValidateNotNullOrEmpty()]
		[ValidatePattern("[0-9]")]
		[int] $Length = 12,

		#When specified, only uses alphanumeric characters for the password
		[Parameter(ValueFromPipeline = $False, ValueFromPipelineByPropertyName = $True)]
		[switch] $ExcludeSymbols
	)

	PROCESS {
		$pw = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
		if (!$ExcludeSymbols) {
    	    $pw += "!$%^-_<>#&@~"
		}
		$password = -join ([Char[]]$pw | Get-Random -count $length)
		Write-Output "`nPassword: $password`n"

		ForEach ($character in [char[]]"$password"){
			[string]$ThisLetter = $character
			switch ($ThisLetter)	{
				a	{$ThisWord = "alpha"}
				b	{$ThisWord = "bravo"}
				c	{$ThisWord = "charlie"}
				d	{$ThisWord = "delta"}
				e	{$ThisWord = "echo"}
				f	{$ThisWord = "foxtrot"}
				g	{$ThisWord = "golf"}
				h	{$ThisWord = "hotel"}
				i	{$ThisWord = "india"}
				j	{$ThisWord = "juliett"}
				k	{$ThisWord = "kilo"}
				l	{$ThisWord = "lima"}
				m	{$ThisWord = "mike"}
				n	{$ThisWord = "november"}
				o	{$ThisWord = "oscar"}
				p	{$ThisWord = "papa"}
				q	{$ThisWord = "quebec"}
				r	{$ThisWord = "romeo"}
				s	{$ThisWord = "sierra"}
				t	{$ThisWord = "tango"}
				u	{$ThisWord = "uniform"}
				v	{$ThisWord = "victor"}
				w	{$ThisWord = "whiskey"}
				x	{$ThisWord = "xray"}
				y	{$ThisWord = "yankee"}
				z	{$ThisWord = "zulu"}
				1	{$ThisWord = "one"}
				2	{$ThisWord = "two"}
				3	{$ThisWord = "three"}
				4	{$ThisWord = "four"}
				5	{$ThisWord = "five"}
				6	{$ThisWord = "six"}
				7	{$ThisWord = "seven"}
				8	{$ThisWord = "eight"}
				9	{$ThisWord = "nine"}
				0	{$ThisWord = "zero"}
				!	{$ThisWord = "exclamation"}
				$	{$ThisWord = "dollar"}
				%	{$ThisWord = "percent"}
				^	{$ThisWord = "carat"}
				-	{$ThisWord = "hyphen"}
				_	{$ThisWord = "underscore"}
				:	{$ThisWord = "colon"}
				`;	{$ThisWord = "semicolon"}
				`{	{$ThisWord = "left-brace"}
				`}	{$ThisWord = "right-brace"}
				`/	{$ThisWord = "backslash"}
				`<	{$ThisWord = "less-than"}
				`>	{$ThisWord = "greater-than"}
				`#	{$ThisWord = "pound"}
				`&  {$ThisWord = "ampersand"}
				`@  {$ThisWord = "at"}
	      			`]  {$ThisWord = "right-bracket"}
	      			`~  {$ThisWord = "tilde"}
	      			default {$ThisWord = "space"}
			}
			if ($ThisLetter -cmatch $ThisLetter.ToUpper()){
				$ThisWord = $ThisWord.ToUpper()
			}
			$phonetic = $phonetic+" " +$ThisWord
		}
		$phonetic = $phonetic.trim()
		Write-Output "Phonetic: $phonetic"
		"Password: $password`nPhonetic: $phonetic" | clip
		Write-Output "`n`nThis information has been sent to the clipboard"
	}
	END{
		Remove-Variable ThisWord
		Remove-Variable ThisLetter
		Remove-Variable Password
		Remove-Variable Phonetic
		Remove-Variable Pw
	}
} # end function New-Password

[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$title = 'Password Length'
$msg   = 'Enter Password Length:'

$Length = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)


$wshell = New-Object -ComObject Wscript.Shell
#$Length = Read-Host "Password Length"
$PopPass = new-Password $Length
#$wshell.Popup("$PopPass")
$PopPass | Out-Gridview -passthru

