Attribute VB_Name = "Module2"
Public Sub SaveAttachmentsToDisk(MItem As MailItem)

    Dim oAttachment As Attachment
    Dim sSaveFolder As String
 
    Dim sndrEmailAdd As String
    Dim sndrEmailRight As String
    Dim sndrEmailPreDot As String
    
    Dim saveName As String

sSaveFolder = "C:\Users\JHARRIS2\Documents\2. Mortgages\Elasticity Model\elasticity-analysis-main\data-raw\"

For Each oAttachment In MItem.Attachments
saveName = sSaveFolder & Format(MItem.ReceivedTime, "yyyymmdd") & "_" & oAttachment.DisplayName
        oAttachment.SaveAsFile saveName
Next
End Sub
