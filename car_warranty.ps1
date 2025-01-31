# Vehicle Extended Warranty Prank Speech Script
Add-Type -AssemblyName System.speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
$speak.Speak("We have been trying to reach you concerning your vehicle's extended warranty. 
You should have received a notice in the mail about your car's extended warranty eligibility. 
Since we have not gotten a response, we are giving you a final courtesy call before we close out your file. 
Press 2 to be removed and placed on our do-not-call list. 
To speak to someone about possibly extending or reinstating your vehicle's warranty, 
press 1 to speak with a warranty specialist.")
