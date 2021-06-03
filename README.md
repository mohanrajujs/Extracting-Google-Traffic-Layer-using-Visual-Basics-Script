# Extracting-Google-Traffic-Layer-using-Visual-Basics-Script
## About
This whole exercise helps one to scrape tarffic layer using _VB script_, a script is designed to extract tiles individually and later whole tiles are stitched together using tile stitching plugin developed by google.
## Step by step method shows how to 1.extract single tile layer, 2.automate extraction of tile layer and 3.finally stitch together to form a single image.
### 1. Defining Bounding Box (Choosing the area)
* #### 1.1 How to get Boundary Information?
![Picture2](https://user-images.githubusercontent.com/5634888/120700143-e937fe00-c4ce-11eb-91d9-48a4d1ada053.png)
####
Go to Google Maps or Google earth. Go to respective area place marker and get the Lat-Long information. Similarly repeat the process for all the remaining edges. (Note: - Make sure Zoom level should be same while recording Lat-Long).

### 2. Tweaking Traffic URL
* #### 2.1 Getting Tile Information from [maptiler.org](http://www.maptiler.org/google-maps-coordinates-tile-bounds-projection/)
Above is the URL which we can get X and Y and Z values. Here x = Latitude | y = Longitude | Z =Zoom level. For Extracting google traffic data, Use Google (X, Y) values from the above given URL.

* #### 2.2 Below is the example URL’s, for getting traffic data with base layer and without base layer.

![Without Baselayer](https://user-images.githubusercontent.com/5634888/120701729-da524b00-c4d0-11eb-88b4-9a45f1a9ddd0.png) &nbsp; &nbsp; ![With Baselyaer](https://user-images.githubusercontent.com/5634888/120702172-65334580-c4d1-11eb-893e-0abd1d3b1420.png)

 ``` https://mts0.google.com/vt?hl=x-local&gl=IN&src=app& ``` x=2928```&```y=1897```&```z=12```&s=G&lyrs=m@254069031,traffic&apistyle=s.tJ37%7Cs.e%3Al%7Cp.v%3Aoff ```

### 3. Traffic layer Extraction & Automation (tile by tile)
* #### 3.1 Tweaking URL & feeding the URL to a automation script.
The below script collects traffic data for every 15 minutes interval. This can be modified based on user requirement
```
Option Explicit

'Private Declare Sub google_traffic Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
ByVal szURL As String, ByVal szFileName As String, _
ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Dim Ret As Long
Dim TimeToRun

Sub google_traffic()

Dim ws As Worksheet
Dim LastRow As Long, i As Long
Dim strPath As String
Dim destinationDir As String
Dim strDirectory As String
Dim objFolder As Variant
Dim FolderName As Variant
Dim fso As Variant
    
destinationDir = "E:\Data_Google_Traffic\"
Set fso = CreateObject("Scripting.FileSystemObject")
'Const OverwriteExisting = True
strDirectory = destinationDir & Replace(Date, "/", "_ ") & " " & Format(Time, "hh_mm_ss")
Set objFolder = fso.CreateFolder(strDirectory)

'Not needed, timer takes care of this issue
'If Not fso.FolderExists(strDirectory) Then
'   Set objFolder = fso.CreateFolder(strDirectory)
'End If
'Set FolderName = fso

Set ws = Sheets("Google_Trafffic")
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
For i = 2 To LastRow
    strPath = objFolder & "\" & ws.Range("A" & i).Value & ".jpg"
    Ret = URLDownloadToFile(0, ws.Range("B" & i).Value, strPath, 0, 0)
    If Ret = 0 Then
       ws.Range("C" & i).Value = "File successfully downloaded"
    Else
       ws.Range("C" & i).Value = "Unable to download the file"
    End If
Application.Wait Now + TimeValue("00:00:10"):
Next i
Call Main
End Sub
 
Sub Main()

TimeToRun = Now + TimeValue("00:01:00")

    Application.OnTime TimeToRun, "google_Traffic"

End Sub


Sub auto_open()

    Call Main

End Sub

Sub auto_close()

    On Error Resume Next
    Application.OnTime TimeToRun, "google_Traffic", , False
    
End Sub

```
* #### 3.2 Additional Instructions for running Vb Script
  - a.	Since these script includes timer function, dedicated system should be maintained in order to get this script run for days and months.
  - b. The script also generates folder based on the timestamp and all the images are stored with in the folder. For each 15 minutes the script generates folder and images (Traffic data) are collected within that folder.

### 4. Stitching each tile using OpenCV & Python
 - Drop the ```stitcher.py``` file into the root directory of your tile map set, where all the numbered folders are
 - Edit two lines in the file, these are commented - one for the number of folders and one for the number of images in each folder
 - Run the script python ```stitcher.py``` [Link](https://gist.github.com/will-hart/133814e92cf45745e9d1)
 - Your output will eventually appear in ```output.png```

### 5. Converting Raster to CSV using OpenCV Python Library

### 6. Calculating Traffic length using QGIS Distance Function

### 7. Sample Output
 
