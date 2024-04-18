'------------------------------------------------------------------------------
'   MIT License
'
'   Copyright (c) 2024 Docuneering Ltd
'
'   Permission is hereby granted, free of charge, to any person obtaining a copy
'   of this software and associated documentation files (the "Software"), to deal
'   in the Software without restriction, including without limitation the rights
'   to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'   copies of the Software, and to permit persons to whom the Software is
'   furnished to do so, subject to the following conditions:
'   
'   The above copyright notice and this permission notice shall be included in all
'   copies or substantial portions of the Software.
'
'   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'   IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'   FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'   AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'   LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'   OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'   SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit

Dim wshShell
Dim objFSO, objFolder, objFile, objXMLDoc, objXMLNode, objXMLDMAddress, objNode, objDMCode, objIssueInfo, objLanguage, objIssueDate, objTechName, objInfoName, objIdentExtension
Dim strScriptPath
Dim strDmList, dmList, dmListFile
Dim strName, strExt, strNameExt
Dim strXPath
Dim xmlDmCode, strDmCode
Dim extensionProducer, extensionCode
Dim modelIdentCode, systemDiffCode, systemCode, subSystemCode, subSubSystemCode, assyCode, disassyCode, disassyCodeVariant, infoCode, infoCodeVariant, itemLocationCode, learnCode, learnEventCode
Dim xmlIssueInfo
Dim issueNumber, inWork
Dim xmlLanguage
Dim language, country
Dim xmlIssueDate
Dim xmlTechName, strTechName
Dim xmlInfoName, strInfoName
Dim strTitle
Dim strLine

'--------------------------------------
'Create common objects
'--------------------------------------
Set wshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'--------------------------------------
'Populate the strScriptPath variable with the physical path from where this VBScript file is being called from
'--------------------------------------
strScriptPath = wshShell.CurrentDirectory


Call buildDmList(strScriptPath)


Sub buildDmList(strPath)
   
   '---------------------------------------
   Set objFolder = objFSO.GetFolder(strPath)
      
      '---------------------------------------
      strDmList = "<dmList>"
      
      '---------------------------------------
      'For each file found in the selected folder do the following
      For Each objFile In objFolder.Files
         
         '---------------------------------------
         strName    = objFSO.GetBasename(objFile.Name)
         strExt     = objFSO.GetExtensionName(objFile.Name)
         strNameExt = objFile.Name
         
         '---------------------------------------
         'If the Extension is "XML" then do the following
         If UCase(strExt) = "XML" Then
            
            '--------------------------------------
            'Open an XML object and load the XML file into memory for processing
            Set objXMLDoc = CreateObject("Msxml2.DOMDocument.6.0")
               
               '---------------------------------------
               objXMLDoc.Async = False
               objXMLDoc.setProperty "MultipleErrorMessages", True
               objXMLDoc.setProperty "ProhibitDTD", False
               objXMLDoc.validateOnParse  = False
               objXMLDoc.resolveExternals = False
               objXMLDoc.Load strPath & "\" & strNameExt
               
               '---------------------------------------
               'XPath to the <dmaddres> or <dmAddress> element
               strXPath = "/dmodule/idstatus/dmaddres | /dmodule/identAndStatusSection/dmAddress"
               Set objXMLNode = objXMLDoc.selectSingleNode(strXPath)
                  
                  '---------------------------------------
                  If NOT objXMLNode Is Nothing Then
                     
                     '---------------------------------------
                     'Create an XML object
                     Set objXMLDMAddress = CreateObject("Msxml2.DOMDocument.6.0")
                        
                        '---------------------------------------
                        'Load XML data into the "objXMLDMAddress" object
                        objXMLDMAddress.Async = False
                        objXMLDMAddress.loadXML(objXMLNode.xml)
                       
                       '---------------------------------------
                        'Clear DMC variables
                        extensionProducer   = ""
                        extensionCode       = ""
                        modelIdentCode      = ""
                        systemDiffCode      = ""
                        systemCode          = ""
                        subSystemCode       = ""
                        subSubSystemCode    = ""
                        assyCode            = ""
                        disassyCode         = ""
                        disassyCodeVariant  = ""
                        infoCode            = ""
                        infoCodeVariant     = ""
                        itemLocationCode    = ""
                        learnCode           = ""
                        learnEventCode      = ""
                        issueNumber         = ""
                        inWork              = ""
                        language            = ""
                        country             = ""
                        
                        '---------------------------------------------------------------------------------------------------------------------
                        'S1000D 3.0 and below
                        If objXMLNode.nodeName = "dmaddres" Then
                           
                           '---------------------------------------
                           strXPath = "/dmaddres/dmc/avee"
                           Set objDMCode = objXMLDMAddress.selectNodes(strXPath) 'selectNodes
                              '---------------------------------------
                              If NOT objDMCode Is Nothing Then
                                 '---------------------------------------
                                 For Each objNode in objDMCode
                                    '---------------------------------------
                                    xmlDmCode          = objNode.xml
                                    '---------------------------------------
                                    modelIdentCode     = objNode.selectSingleNode("modelic").text
                                    systemDiffCode     = objNode.selectSingleNode("sdc").text
                                    systemCode         = objNode.selectSingleNode("chapnum").text
                                    subSystemCode      = objNode.selectSingleNode("section").text
                                    subSubSystemCode   = objNode.selectSingleNode("subsect").text
                                    assyCode           = objNode.selectSingleNode("subject").text
                                    disassyCode        = objNode.selectSingleNode("discode").text
                                    disassyCodeVariant = objNode.selectSingleNode("discodev").text
                                    infoCode           = objNode.selectSingleNode("incode").text
                                    infoCodeVariant    = objNode.selectSingleNode("incodev").text
                                    itemLocationCode   = objNode.selectSingleNode("itemloc").text
                                 Next
                              End If
                              '---------------------------------------
                           Set objDMCode = Nothing
                           
                           '---------------------------------------
                           strXPath = "/dmaddres/issno"
                           Set objIssueInfo = objXMLDMAddress.selectSingleNode(strXPath) 'selectSingleNode
                              '---------------------------------------
                              'To avoid errors ... as long as it finds the <issno> element
                              If NOT objIssueInfo Is Nothing Then
                                 '---------------------------------------
                                 xmlIssueInfo = objIssueInfo.xml
                                 '---------------------------------------
                                 issueNumber = objIssueInfo.getAttribute("issno")
                                 inWork      = objIssueInfo.getAttribute("inwork")
                                 '---------------------------------------
                              End If
                              '---------------------------------------
                           Set objIssueInfo = Nothing
                           
                           '---------------------------------------
                           strXPath = "/dmaddres/language"
                           Set objLanguage = objXMLDMAddress.selectSingleNode(strXPath) 'selectSingleNode
                              '---------------------------------------
                              'To avoid errors ... as long as it finds the <language> element
                              If NOT objLanguage Is Nothing Then
                                 '---------------------------------------
                                 xmlLanguage = objLanguage.xml
                                 '---------------------------------------
                                 language = objLanguage.getAttribute("language")
                                 country  = objLanguage.getAttribute("country")
                                 '---------------------------------------
                              End If
                              '---------------------------------------
                           Set objLanguage = Nothing
                           
                           '---------------------------------------                        
                           strXPath = "/dmaddres/issdate"
                           Set objIssueDate = objXMLDMAddress.selectSingleNode(strXPath)
                              '---------------------------------------
                              If NOT objIssueDate Is Nothing Then
                                 '---------------------------------------
                                 xmlIssueDate = objIssueDate.xml
                                 '---------------------------------------
                              End If
                              '---------------------------------------
                           Set objIssueDate = Nothing
                           
                           '---------------------------------------
                           strXPath = "/dmaddres/dmtitle/techname"
                           Set objTechName = objXMLDMAddress.selectSingleNode(strXPath) 'selectSingleNode
                              '---------------------------------------
                              'To avoid errors ... as long as it finds the <language> element
                              If NOT objTechName Is Nothing Then
                                 '---------------------------------------
                                 xmlTechName = objTechName.xml
                                 strTechName = objTechName.text
                                 '---------------------------------------
                              End If
                              '---------------------------------------
                           Set objTechName = Nothing
                           
                           '---------------------------------------
                           strXPath = "/dmaddres/dmtitle/infoname"
                           Set objInfoName = objXMLDMAddress.selectSingleNode(strXPath) 'selectSingleNode
                              '---------------------------------------
                              'To avoid errors ... as long as it finds the <language> element
                              If NOT objInfoName Is Nothing Then
                                 '---------------------------------------
                                 xmlInfoName = objInfoName.xml
                                 strInfoName = objInfoName.text
                                 '---------------------------------------
                              End If
                              '---------------------------------------
                           Set objInfoName = Nothing
                           
                        '---------------------------------------------------------------------------------------------------------------------
                        'S1000D 4.x and above
                        Else
                           
                           '---------------------------------------
                           strXPath = "/dmAddress/dmIdent/identExtension"
                           Set objIdentExtension = objXMLDMAddress.selectSingleNode(strXPath)
                              '---------------------------------------
                              If NOT objIdentExtension Is Nothing Then
                                 '---------------------------------------
                                 extensionProducer = objIdentExtension.getAttribute("extensionProducer")
                                 extensionCode     = objIdentExtension.getAttribute("extensionCode")
                                 '---------------------------------------
                              End If
                              '---------------------------------------
                           Set objIdentExtension = Nothing
                           
                           '---------------------------------------
                           strXPath = "/dmAddress/dmIdent/dmCode"
                           Set objDMCode = objXMLDMAddress.selectSingleNode(strXPath)
                              '---------------------------------------
                              If NOT objDMCode Is Nothing Then
                                 '---------------------------------------
                                 xmlDmCode           = objDMCode.xml
                                 '---------------------------------------
                                 modelIdentCode      = objDMCode.getAttribute("modelIdentCode")
                                 systemDiffCode      = objDMCode.getAttribute("systemDiffCode")
                                 systemCode          = objDMCode.getAttribute("systemCode")
                                 subSystemCode       = objDMCode.getAttribute("subSystemCode")
                                 subSubSystemCode    = objDMCode.getAttribute("subSubSystemCode")
                                 assyCode            = objDMCode.getAttribute("assyCode")
                                 disassyCode         = objDMCode.getAttribute("disassyCode")
                                 disassyCodeVariant  = objDMCode.getAttribute("disassyCodeVariant")
                                 infoCode            = objDMCode.getAttribute("infoCode")
                                 infoCodeVariant     = objDMCode.getAttribute("infoCodeVariant")
                                 itemLocationCode    = objDMCode.getAttribute("itemLocationCode")
                                 'Learning doctype
                                 learnCode           = objDMCode.getAttribute("learnCode")
                                 learnEventCode      = objDMCode.getAttribute("learnEventCode")
                                 '---------------------------------------
                              End If
                              '---------------------------------------
                           Set objDMCode = Nothing
                           
                           '---------------------------------------                        
                           strXPath = "/dmAddress/dmIdent/issueInfo"
                           Set objIssueInfo = objXMLDMAddress.selectSingleNode(strXPath)
                              '---------------------------------------
                              If NOT objIssueInfo Is Nothing Then
                                 '---------------------------------------
                                 xmlIssueInfo = objIssueInfo.xml
                                 '---------------------------------------
                                 issueNumber  = objIssueInfo.getAttribute("issueNumber")
                                 inWork       = objIssueInfo.getAttribute("inWork")
                                 '---------------------------------------
                              End If
                              '---------------------------------------
                           Set objIssueInfo = Nothing
                           
                           '---------------------------------------                        
                           strXPath = "/dmAddress/dmIdent/language"
                           Set objLanguage = objXMLDMAddress.selectSingleNode(strXPath)
                              '---------------------------------------
                              If NOT objLanguage Is Nothing Then
                                 '---------------------------------------
                                 xmlLanguage = objLanguage.xml
                                 '---------------------------------------
                                 language    = objLanguage.getAttribute("languageIsoCode")
                                 country     = objLanguage.getAttribute("countryIsoCode")
                                 '---------------------------------------
                              End If
                              '---------------------------------------
                           Set objLanguage = Nothing
                           
                           '---------------------------------------                        
                           strXPath = "/dmAddress/dmAddressItems/issueDate"
                           Set objIssueDate = objXMLDMAddress.selectSingleNode(strXPath)
                              '---------------------------------------
                              If NOT objIssueDate Is Nothing Then
                                 '---------------------------------------
                                 xmlIssueDate = objIssueDate.xml
                                 '---------------------------------------
                              End If
                              '---------------------------------------
                           Set objIssueDate = Nothing
                           
                           '---------------------------------------                        
                           strXPath = "/dmAddress/dmAddressItems/dmTitle/techName"
                           Set objTechName = objXMLDMAddress.selectSingleNode(strXPath)
                              '---------------------------------------
                              If NOT objTechName Is Nothing Then
                                 '---------------------------------------
                                 xmlTechName = objTechName.xml
                                 strTechName = objTechName.text
                                 '---------------------------------------
                              End If
                              '---------------------------------------
                           Set objTechName = Nothing
                           
                           '---------------------------------------                        
                           strXPath = "/dmAddress/dmAddressItems/dmTitle/infoName"
                           Set objInfoName = objXMLDMAddress.selectSingleNode(strXPath)
                              '---------------------------------------
                              If NOT objInfoName Is Nothing Then
                                 '---------------------------------------
                                 xmlInfoName = objInfoName.xml
                                 strInfoName = objInfoName.text
                                 '---------------------------------------
                              End If
                              '---------------------------------------
                           Set objInfoName = Nothing
                           
                        End If
                        
                        '---------------------------------------
                        If NOT extensionProducer = "" AND NOT extensionCode = "" Then
                           strDmCode = "DME-"
                           strDmCode = strDmCode & extensionProducer & "-" & extensionCode & "-"
                        Else
                           strDmCode = "DMC-"
                        End If
                        
                        '---------------------------------------
                        strDmCode = strDmCode & modelIdentCode & "-" & systemDiffCode & "-" & systemCode & "-" & subSystemCode & subSubSystemCode & "-" & assyCode & "-" & disassyCode & disassyCodeVariant & "-" & infoCode & infoCodeVariant & "-" & itemLocationCode
                        
                        '---------------------------------------
                        If NOT learnCode = "" AND NOT learnEventCode = "" Then
                           strDmCode = strDmCode & "-" & learnCode & learnEventCode
                        End If
                        
                        '---------------------------------------
                        If NOT issueNumber = "" AND NOT inWork = "" Then
                           strDmCode = strDmCode & "_" & issueNumber & "-" & inWork
                        End If
                        
                        '---------------------------------------
                        If NOT language = "" AND NOT country = "" Then
                           strDmCode = strDmCode & "_" & language & "-" & country
                        End If
                        
                        '---------------------------------------
                        strTitle = "[" & strDmCode & "] " & strTechName & " - " & strInfoName
                        
                        '---------------------------------------------------------------------------------------------------------------------
                        'S1000D 3.0 and below
                        If objXMLNode.nodeName = "dmaddres" Then
                           '---------------------------------------
                           strLine = "<dmodule>"
                           strLine = strLine & "<title>" & strTitle & "</title>"
                           strLine = strLine & "<refdm>"
                           strLine = strLine & "<dmc>"
                           
                           strLine = strLine & xmlDmCode
                           
                           strLine = strLine & "</dmc>"
                           strLine = strLine & "<dmtitle>"
                           
                           strLine = strLine & xmlTechName
                           strLine = strLine & xmlInfoName
                           
                           strLine = strLine & "</dmtitle>"
                           
                           strLine = strLine & xmlIssueInfo
                           strLine = strLine & xmlIssueDate
                           strLine = strLine & xmlLanguage
                           
                           strLine = strLine & "</refdm>"
                           strLine = strLine & "</dmodule>"
                           
                        '---------------------------------------------------------------------------------------------------------------------
                        'S1000D 4.x and above
                        Else
                           '---------------------------------------
                           strLine = "<dmodule>"
                           strLine = strLine & "<title>" & strTitle & "</title>"
                           strLine = strLine & "<dmRef>"
                           strLine = strLine & "<dmRefIdent>"
                           
                           strLine = strLine & xmlDmCode
                           strLine = strLine & xmlIssueInfo
                           strLine = strLine & xmlLanguage
                           
                           strLine = strLine & "</dmRefIdent>"
                           strLine = strLine & "<dmRefAddressItems>"
                           strLine = strLine & "<dmTitle>"
                           
                           strLine = strLine & xmlTechName
                           strLine = strLine & xmlInfoName
                           
                           strLine = strLine & "</dmTitle>"
                           
                           strLine = strLine & xmlIssueDate
                           
                           strLine = strLine & "</dmRefAddressItems>"
                           strLine = strLine & "</dmRef>"
                           strLine = strLine & "</dmodule>"
                           
                        End If
                        '---------------------------------------
                        strDmList = strDmList & strLine
                        '---------------------------------------
                     Set objXMLDMAddress = Nothing
                     '---------------------------------------
                  End If
               '---------------------------------------
               Set objXMLNode = Nothing
               '---------------------------------------
            Set objXMLDoc = Nothing
            '---------------------------------------
         End If
         '---------------------------------------
      Next
      '---------------------------------------
      strDmList = strDmList & "</dmList>"
      
      '---------------------------------------
      'Write dmList to the "dmList.xml" file
      Set dmListFile = objFSO.OpenTextFile("dmList.xml" ,2 , True)
      dmListFile.WriteLine strDmList
      
      '---------------------------------------
   Set objFSO = Nothing
   Set objFolder = Nothing
   
   '---------------------------------------
End Sub
'---------------------------------------

MsgBox "...build complete"
