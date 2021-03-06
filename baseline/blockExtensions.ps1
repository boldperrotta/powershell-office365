$ruleName = "Block Uncommon Attachment Extensions"

$rule = Get-TransportRule | Where-Object {$_.Identity -contains $ruleName}
 
if (!$rule) {
    Write-Host "Rule not found, creating rule" -ForegroundColor Green
    New-TransportRule -Name $ruleName -Priority 0 -FromScope "NotInOrganization" -AttachmentExtensionMatchesWords .adp,.app,.asp,.bas,.bat,.cer,.chm,.cmd,.cnt,.com,.cpl,.crt,.csh,.der,.exe,.fxp,.gadget,.hlp,.hpj,.hta,.inf,.ins,.isp,.its,.js,.jse,.ksh,.lnk,.mad,.maf,.mag,.mam,.maq,.mar,.mas,.mat,.mau,.mav,.maw,.mda,.mdb,.mde,.mdt,.mdw,.mdz,.msc,.msh,.msh1,.msh2,.mshxml,.msh1xml,.msh2xml,.msi,.msp,.mst,.ops,.osd,.pcd,.pif,.plg,.prf,.prg,.pst,.reg,.scf,.scr,.sct,.shb,.shs,.ps1,.ps1xml,.ps2,.ps2xml,.psc1,.psc2,.tmp,.url,.vb,.vbe,.vbp,.vbs,.vsmacros,.vsw,.ws,.wsc,.wsf,.wsh,.xnk,.ade,.cla,.class,.grp,.jar,.mcf,.ocx,.pl,.xbap -DeleteMessage $true
       }
else {
    Write-Host "Rule found, updating rule" -ForegroundColor Green
    Set-TransportRule -Identity $ruleName -Priority 0 -FromScope "NotInOrganization" -AttachmentExtensionMatchesWords .adp,.app,.asp,.bas,.bat,.cer,.chm,.cmd,.cnt,.com,.cpl,.crt,.csh,.der,.exe,.fxp,.gadget,.hlp,.hpj,.hta,.inf,.ins,.isp,.its,.js,.jse,.ksh,.lnk,.mad,.maf,.mag,.mam,.maq,.mar,.mas,.mat,.mau,.mav,.maw,.mda,.mdb,.mde,.mdt,.mdw,.mdz,.msc,.msh,.msh1,.msh2,.mshxml,.msh1xml,.msh2xml,.msi,.msp,.mst,.ops,.osd,.pcd,.pif,.plg,.prf,.prg,.pst,.reg,.scf,.scr,.sct,.shb,.shs,.ps1,.ps1xml,.ps2,.ps2xml,.psc1,.psc2,.tmp,.url,.vb,.vbe,.vbp,.vbs,.vsmacros,.vsw,.ws,.wsc,.wsf,.wsh,.xnk,.ade,.cla,.class,.grp,.jar,.mcf,.ocx,.pl,.xbap -DeleteMessage $true
}
