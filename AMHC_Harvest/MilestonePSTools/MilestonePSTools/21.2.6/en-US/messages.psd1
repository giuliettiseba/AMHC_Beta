
@{
    # This is a table of en-US messages used in the functions in this module. We should
    # work on getting ALL string literals in Write-* cmdlets moved into this file so that
    # we can localize this file along with Get-Help content, and automatically improve
    # the messaging for languages we end up offering localized support for.

    # Look for Import-LocalizedData in the module manifest where this file will be imported
    # into a script-scope variable. Functions can then use these messages by referencing
    # $script:Messages.ListingAllRecorders for example.

    MustBeAdminToReadPasswords = 'You must be in the Administrators role to read hardware passwords.'
    ListingAllRecorders = 'Listing all recording servers'
    CallingGetItemState = 'Calling Get-ItemState'
    StartingFillChildrenThreadJob = 'Starting FillChildren threadjob'
    CameraOnSiteWithIdNotFound = 'Camera not found on site {0} with ID {1}.'
    NotConnectedToAManagementServer = 'Not connected to a Milestone XProtect Management Server.'
    PrebufferSecondsExceedsMaximumValue = 'PrebufferSeconds exceeds the maximum of value for in-memory buffering. The value will be updated to 15 seconds.'
    ClientServiceValidateResult = 'Validation error: Failed to set {0} to "{1}". Reason: {2}.'
}

# SIG # Begin signature block
# MIId+QYJKoZIhvcNAQcCoIId6jCCHeYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUJar5lxM3FljWokv5vGeVFnX/
# H0+gghglMIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMB4XDTIxMDEwMTAwMDAwMFoXDTMxMDEw
# NjAwMDAwMFowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAMLmYYRnxYr1DQikRcpja1HXOhFCvQp1dU2UtAxQ
# tSYQ/h3Ib5FrDJbnGlxI70Tlv5thzRWRYlq4/2cLnGP9NmqB+in43Stwhd4CGPN4
# bbx9+cdtCT2+anaH6Yq9+IRdHnbJ5MZ2djpT0dHTWjaPxqPhLxs6t2HWc+xObTOK
# fF1FLUuxUOZBOjdWhtyTI433UCXoZObd048vV7WHIOsOjizVI9r0TXhG4wODMSlK
# XAwxikqMiMX3MFr5FK8VX2xDSQn9JiNT9o1j6BqrW7EdMMKbaYK02/xWVLwfoYer
# vnpbCiAvSwnJlaeNsvrWY4tOpXIc7p96AXP4Gdb+DUmEvQECAwEAAaOCAbgwggG0
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEEGA1UdIAQ6MDgwNgYJYIZIAYb9bAcBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAfBgNVHSMEGDAWgBT0tuEgHf4prtLk
# YaWyoiWyyBc1bjAdBgNVHQ4EFgQUNkSGjqS6sGa+vCgtHUQ23eNqerwwcQYDVR0f
# BGowaDAyoDCgLoYsaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC10cy5jcmwwMqAwoC6GLGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtdHMuY3JsMIGFBggrBgEFBQcBAQR5MHcwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBPBggrBgEFBQcwAoZDaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRFRpbWVzdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEASBzctemaI7znGucgDo5nRv1CclF0CiNH
# o6uS0iXEcFm+FKDlJ4GlTRQVGQd58NEEw4bZO73+RAJmTe1ppA/2uHDPYuj1UUp4
# eTZ6J7fz51Kfk6ftQ55757TdQSKJ+4eiRgNO/PT+t2R3Y18jUmmDgvoaU+2QzI2h
# F3MN9PNlOXBL85zWenvaDLw9MtAby/Vh/HUIAHa8gQ74wOFcz8QRcucbZEnYIpp1
# FUL1LTI4gdr0YKK6tFL7XOBhJCVPst/JKahzQ1HavWPWH1ub9y4bTxMd90oNcX6X
# t/Q/hOvB46NJofrOp79Wz7pZdmGJX36ntI5nePk2mOHLKNpbh6aKLzCCBTEwggQZ
# oAMCAQICEAqhJdbWMht+QeQF2jaXwhUwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTE2MDEwNzEyMDAwMFoXDTMxMDEwNzEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAL3QMu5LzY9/3am6gpnF
# OVQoV7YjSsQOB0UzURB90Pl9TWh+57ag9I2ziOSXv2MhkJi/E7xX08PhfgjWahQA
# OPcuHjvuzKb2Mln+X2U/4Jvr40ZHBhpVfgsnfsCi9aDg3iI/Dv9+lfvzo7oiPhis
# EeTwmQNtO4V8CdPuXciaC1TjqAlxa+DPIhAPdc9xck4Krd9AOly3UeGheRTGTSQj
# MF287DxgaqwvB8z98OpH2YhQXv1mblZhJymJhFHmgudGUP2UKiyn5HU+upgPhH+f
# MRTWrdXyZMt7HgXQhBlyF/EXBu89zdZN7wZC/aJTKk+FHcQdPK/P2qwQ9d2srOlW
# /5MCAwEAAaOCAc4wggHKMB0GA1UdDgQWBBT0tuEgHf4prtLkYaWyoiWyyBc1bjAf
# BgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzASBgNVHRMBAf8ECDAGAQH/
# AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB5BggrBgEF
# BQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBD
# BggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDig
# NoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDBQBgNVHSAESTBHMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYc
# aHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzALBglghkgBhv1sBwEwDQYJKoZI
# hvcNAQELBQADggEBAHGVEulRh1Zpze/d2nyqY3qzeM8GN0CE70uEv8rPAwL9xafD
# DiBCLK938ysfDCFaKrcFNB1qrpn4J6JmvwmqYN92pDqTD/iy0dh8GWLoXoIlHsS6
# HHssIeLWWywUNUMEaLLbdQLgcseY1jxk5R9IEBhfiThhTWJGJIdjjJFSLK8pieV4
# H9YLFKWA1xJHcLN11ZOFk362kmf7U2GJqPVrlsD0WGkNfMgBsbkodbeZY4UijGHK
# eZR+WfyMD+NvtQEmtmyl7odRIeRYYJu6DC0rbaLEfrvEJStHAgh8Sa4TtuF8QkIo
# xhhWz0E0tmZdtnR79VYzIi8iNrJLokqV2PWmjlIwggbmMIIEzqADAgECAhB3vQ4D
# obcI+FSrBnIQ2QRHMA0GCSqGSIb3DQEBCwUAMFMxCzAJBgNVBAYTAkJFMRkwFwYD
# VQQKExBHbG9iYWxTaWduIG52LXNhMSkwJwYDVQQDEyBHbG9iYWxTaWduIENvZGUg
# U2lnbmluZyBSb290IFI0NTAeFw0yMDA3MjgwMDAwMDBaFw0zMDA3MjgwMDAwMDBa
# MFkxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMS8wLQYD
# VQQDEyZHbG9iYWxTaWduIEdDQyBSNDUgQ29kZVNpZ25pbmcgQ0EgMjAyMDCCAiIw
# DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBANZCTfnjT8Yj9GwdgaYw90g9z9Dl
# jeUgIpYHRDVdBs8PHXBg5iZU+lMjYAKoXwIC947Jbj2peAW9jvVPGSSZfM8RFpsf
# e2vSo3toZXer2LEsP9NyBjJcW6xQZywlTVYGNvzBYkx9fYYWlZpdVLpQ0LB/okQZ
# 6dZubD4Twp8R1F80W1FoMWMK+FvQ3rpZXzGviWg4QD4I6FNnTmO2IY7v3Y2FQVWe
# HLw33JWgxHGnHxulSW4KIFl+iaNYFZcAJWnf3sJqUGVOU/troZ8YHooOX1ReveBb
# z/IMBNLeCKEQJvey83ouwo6WwT/Opdr0WSiMN2WhMZYLjqR2dxVJhGaCJedDCndS
# sZlRQv+hst2c0twY2cGGqUAdQZdihryo/6LHYxcG/WZ6NpQBIIl4H5D0e6lSTmpP
# VAYqgK+ex1BC+mUK4wH0sW6sDqjjgRmoOMieAyiGpHSnR5V+cloqexVqHMRp5rC+
# QBmZy9J9VU4inBDgoVvDsy56i8Te8UsfjCh5MEV/bBO2PSz/LUqKKuwoDy3K1JyY
# ikptWjYsL9+6y+JBSgh3GIitNWGUEvOkcuvuNp6nUSeRPPeiGsz8h+WX4VGHaeki
# zIPAtw9FbAfhQ0/UjErOz2OxtaQQevkNDCiwazT+IWgnb+z4+iaEW3VCzYkmeVmd
# a6tjcWKQJQ0IIPH/AgMBAAGjggGuMIIBqjAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0l
# BAwwCgYIKwYBBQUHAwMwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU2rON
# wCSQo2t30wygWd0hZ2R2C3gwHwYDVR0jBBgwFoAUHwC/RoAK/Hg5t6W0Q9lWULvO
# ljswgZMGCCsGAQUFBwEBBIGGMIGDMDkGCCsGAQUFBzABhi1odHRwOi8vb2NzcC5n
# bG9iYWxzaWduLmNvbS9jb2Rlc2lnbmluZ3Jvb3RyNDUwRgYIKwYBBQUHMAKGOmh0
# dHA6Ly9zZWN1cmUuZ2xvYmFsc2lnbi5jb20vY2FjZXJ0L2NvZGVzaWduaW5ncm9v
# dHI0NS5jcnQwQQYDVR0fBDowODA2oDSgMoYwaHR0cDovL2NybC5nbG9iYWxzaWdu
# LmNvbS9jb2Rlc2lnbmluZ3Jvb3RyNDUuY3JsMFYGA1UdIARPME0wQQYJKwYBBAGg
# MgEyMDQwMgYIKwYBBQUHAgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29tL3Jl
# cG9zaXRvcnkvMAgGBmeBDAEEATANBgkqhkiG9w0BAQsFAAOCAgEACIhyJsav+qxf
# BsCqjJDa0LLAopf/bhMyFlT9PvQwEZ+PmPmbUt3yohbu2XiVppp8YbgEtfjry/Rh
# ETP2ZSW3EUKL2Glux/+VtIFDqX6uv4LWTcwRo4NxahBeGQWn52x/VvSoXMNOCa1Z
# a7j5fqUuuPzeDsKg+7AE1BMbxyepuaotMTvPRkyd60zsvC6c8YejfzhpX0FAZ/ZT
# fepB7449+6nUEThG3zzr9s0ivRPN8OHm5TOgvjzkeNUbzCDyMHOwIhz2hNabXAAC
# 4ShSS/8SS0Dq7rAaBgaehObn8NuERvtz2StCtslXNMcWwKbrIbmqDvf+28rrvBfL
# uGfr4z5P26mUhmRVyQkKwNkEcUoRS1pkw7x4eK1MRyZlB5nVzTZgoTNTs/Z7KtWJ
# QDxxpav4mVn945uSS90FvQsMeAYrz1PYvRKaWyeGhT+RvuB4gHNU36cdZytqtq5N
# iYAkCFJwUPMB/0SuL5rg4UkI4eFb1zjRngqKnZQnm8qjudviNmrjb7lYYuA2eDYB
# +sGniXomU6Ncu9Ky64rLYwgv/h7zViniNZvY/+mlvW1LWSyJLC9Su7UpkNpDR7xy
# 3bzZv4DB3LCrtEsdWDY3ZOub4YUXmimi/eYI0pL/oPh84emn0TCOXyZQK8ei4pd3
# iu/YTT4m65lAYPM8Zwy2CHIpNVOBNNwwggcAMIIE6KADAgECAgxVg/P+wUlFDmSx
# M2wwDQYJKoZIhvcNAQELBQAwWTELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2Jh
# bFNpZ24gbnYtc2ExLzAtBgNVBAMTJkdsb2JhbFNpZ24gR0NDIFI0NSBDb2RlU2ln
# bmluZyBDQSAyMDIwMB4XDTIxMTEyOTE1NDEwMVoXDTIyMTEzMDE1NDEwMVowbjEL
# MAkGA1UEBhMCREsxETAPBgNVBAcMCEJyw7huZGJ5MR4wHAYDVQQKExVNaWxlc3Rv
# bmUgU3lzdGVtcyBBL1MxDDAKBgNVBAsMA1ImRDEeMBwGA1UEAxMVTWlsZXN0b25l
# IFN5c3RlbXMgQS9TMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA0Eti
# hqPkmu1KV6LRSN9xz96UtyFGEZhOdBFPstJsSKMXQRCYcH3wcVxz6/pERdOTllrU
# uojlhmJKXUJ04ak7aHxY6/5WzUQZmm5b5uoZ27p9qSOVVgkYnpgwPAs41b5bC8qB
# V+NSELQniXQEOuacIRHN+oynZitag8Fy/7qtDzYaqmH74PBABr7vBMOUovEuxAxa
# r6v0dRIc2MYbWqOfF6jTJTX9fG0hW4nGQsC04EbENAdMHCZubiTswL20FgbjNq9O
# LpCp+Eu7sYRkmnHz2Kxd++9QlS4DtmI6Hbw6jy7a2WQP0vAsrqdVd1nh2DCbf058
# rHUTXNT/csXv82uSpwdLpSZVZigaKFsTmBmW94sVFK+TQYTe/4lpo+F/sMyPrw0i
# Fv/jJMlp5e3WKa8leiAIFiEfu7vnmI1FFslOlVfHYHc1fe2ERCe1DDW/hq3KFz8D
# 1q2CGMJpY6zY2iZ/mq2bJnMRASM9qOtRdTYLxPzm697bUdgK7p8SLtm1TzbzS1Js
# XpgRslxWXUWUAkSUeEeMXZHaF3wXZIn507FD/oupj0Goc4riHPxjhm9avY5tcMGY
# 8pflyYG1OOjNUlcHhW/cFX3/Tzr4UB2/sWkJ1Jopm22ZopCoDf905LrmgZOlVXc8
# cgApylcnpUKN9bl9XfqXxhYiaw0nz4hfImpjkTsCAwEAAaOCAbEwggGtMA4GA1Ud
# DwEB/wQEAwIHgDCBmwYIKwYBBQUHAQEEgY4wgYswSgYIKwYBBQUHMAKGPmh0dHA6
# Ly9zZWN1cmUuZ2xvYmFsc2lnbi5jb20vY2FjZXJ0L2dzZ2NjcjQ1Y29kZXNpZ25j
# YTIwMjAuY3J0MD0GCCsGAQUFBzABhjFodHRwOi8vb2NzcC5nbG9iYWxzaWduLmNv
# bS9nc2djY3I0NWNvZGVzaWduY2EyMDIwMFYGA1UdIARPME0wQQYJKwYBBAGgMgEy
# MDQwMgYIKwYBBQUHAgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29tL3JlcG9z
# aXRvcnkvMAgGBmeBDAEEATAJBgNVHRMEAjAAMEUGA1UdHwQ+MDwwOqA4oDaGNGh0
# dHA6Ly9jcmwuZ2xvYmFsc2lnbi5jb20vZ3NnY2NyNDVjb2Rlc2lnbmNhMjAyMC5j
# cmwwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHwYDVR0jBBgwFoAU2rONwCSQo2t30wyg
# Wd0hZ2R2C3gwHQYDVR0OBBYEFIdg9gMxxfcpe7gViT9QiXomUUGjMA0GCSqGSIb3
# DQEBCwUAA4ICAQAY8MM+XlsN3AEj6VASoSGTEVEIdMSZD0p2UPa8zkTWm7ZX0JLN
# O81eTCDRXzdS9jAz7U+ldJZwKYeogKshoBkPN0jBz+BCQNhOBYNfpGCFdYwH/cIF
# RPbcPB2sIQpe9Lr4+ZrU7MT5kx2ltznSaLHDf4wwvow0FsdrWfqJhNMsg4eYAuXU
# xnq0dmG2eJzB8XzoiFNhfv215Z45zYlG2vlczZnV2H8VfgvGA1Y6zLE7hLn5xaWg
# s+lwp2e87KIshP8qd+0DK3H+g0pFdcSQKjNnEgoASBAsdxSI0OLbZOLElOqAuaD2
# RAeEleEm0Ww1jZrncpDr8bq4LqT1q+fBMdPOhU28L+V20lHJrLwpjSWtatnKzEn2
# C5pBMgO2tcjPf7HH/Q30HIgCN6e9CxTxT/oO3eyxl7E9KtBL3edjE2/qFhpLVbdo
# P0KUb3go5jsnF9nmkualdsFICtzI7q4QARQi9+toixHk53eCUnxYEo9E01L45y6E
# BrvUWjvdj/A4RQZNjaiokzMYRluj6detWnN1cP0S18z3gwDlKA61bEAdpqvo3ifE
# oDtsCVQQty0duVp8nPzh8jT12V9+nQ9WQfch8x8TEm1qyNpHsIYDKGwsSLju95Zq
# ApYO7GFvzxvRjZMnUw85ePn/n8BY6HhALbPUmkn4PaXEee36XtllOAVvozGCBT4w
# ggU6AgEBMGkwWTELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYt
# c2ExLzAtBgNVBAMTJkdsb2JhbFNpZ24gR0NDIFI0NSBDb2RlU2lnbmluZyBDQSAy
# MDIwAgxVg/P+wUlFDmSxM2wwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFM+s+aXbwh1nV3fIt4WM
# uFnyIJ0jMA0GCSqGSIb3DQEBAQUABIICAJ70wEAYugaXNlNJMzK3Ei0hCve3VTVu
# NFkiYZwPD8AfcHSyNThNHo4PgWmKgjKpJr0mIp5xsyZsl7NYtoLGeRXwGxBBaFrx
# JJ/x+L/8oMda8N4925MsBMwGVNLthVFQOn3wKn5fovuN1hviEN5FonjZvtkHiwLw
# 35zanFEVvnTrFZ/WyqHmYIpE3AXR9hTqoBZPFAVAGBeaFdKi9G1iJC4I1dOgGPXX
# H6jDvsBvg/giAMdeSveJaSvM7G7bc666/WY3huuyOCgPp027E0F1RdNd3rZ/thXd
# 43xvaTIG9XZt2EeOEWmpYkdqe+r+GsUpXZ8k15P6e7LID7CfgeIN0NmzFR2DUwXK
# dW6kYRruIltOuiGsgvOTm7Wpr3HbjB8AhimhTBjxdH3Q/7/DdU4Vv79mmQ49eX+E
# OgQc9MhydBQKrVGRr2xCqnMkjZGdEirEGbBSbTSLQ0pY1K0f6OZoc/oOYKpaBPjd
# 0fKeLLxvbk0rvgPg8cSROTD5JIHIKJMho6h4Pb/HFvsLuwEZNsFAfauwyI5Mxp96
# UdkhCcoQjPZhGsEQcCAk0lWIFOMohPIBSJDHcfxNYlFjpCS+59Q+gmJJOVJUihlC
# kH4fxu/5cyJmKeEHAVLOu+K5WhShW/GLiPY3QRPwVIjSNb6U4QmPOVHp6bcNW8g9
# 7UzTVr2GbbxJoYICMDCCAiwGCSqGSIb3DQEJBjGCAh0wggIZAgEBMIGGMHIxCzAJ
# BgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5k
# aWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBU
# aW1lc3RhbXBpbmcgQ0ECEA1CSuC+Ooj/YEAhzhQA8N0wDQYJYIZIAWUDBAIBBQCg
# aTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yMjAy
# MDkwNjMxNDNaMC8GCSqGSIb3DQEJBDEiBCCVq0Q8JXOQDbb5KgbPZ6okfjbDt1qn
# Z0YbgVf3/ISI7DANBgkqhkiG9w0BAQEFAASCAQCaqZjOs7OVM6jrYkAS0ZIWTrAs
# +efkZHALYVzr/1WCLqvbFFQsm1CHczqfQk3JHYCdXiPM0QyQPAqiM7W2fBdXRFR4
# orwmoSf3cXI+ab8hw8NRfjN6tihCk9+dR+sRluhE0OamDjqLo9Ghmwe42Nemr+uP
# CYXPX/3zUxCUs5Vltxz+LX59LuznA2ddYwqYMFQo25X4EvN93Tg+KfpWefDxjw2f
# 2+8tXgYIcjpIHMb3NCt0PWNPBfW7j5oFUDylFj/Z30jUcwax1Yr0UQ69ZZ5DhJZP
# 3dMNjqOIL3csw5O56pNWf4c/HWnSla4dLoXnzSC2/XmUPlFlzzuElmUt+qBO
# SIG # End signature block
