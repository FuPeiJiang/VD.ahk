; https://www.autohotkey.com/boards/viewtopic.php?t=72797
; MIDL_INTERFACE("BC042D67-9047-33F6-881B-6746C1B218B8")
; IFaceDetectorStatics : public IInspectable

msgbox % facedetect("face.jpg")
ExitApp



facedetect(file, maxheight := 2000)
{
   static BitmapDecoderStatics, BitmapEncoderStatics, SoftwareBitmapStatics, FaceDetector, SupportedBitmapPixelFormats
   if (FaceDetector = "")
   {
      CreateClass("Windows.Graphics.Imaging.BitmapDecoder", IBitmapDecoderStatics := "{438CCB26-BCEF-4E95-BAD6-23A822E58D01}", BitmapDecoderStatics)
      CreateClass("Windows.Graphics.Imaging.BitmapEncoder", IBitmapEncoderStatics := "{A74356A7-A4E4-4EB9-8E40-564DE7E1CCB2}", BitmapEncoderStatics)
      CreateClass("Windows.Graphics.Imaging.SoftwareBitmap", ISoftwareBitmapStatics := "{DF0385DB-672F-4A9D-806E-C2442F343E86}", SoftwareBitmapStatics)
      CreateClass("Windows.Media.FaceAnalysis.FaceDetector", IFaceDetectorStatics := "{BC042D67-9047-33F6-881B-6746C1B218B8}", FaceDetectorStatics)
      DllCall(NumGet(NumGet(FaceDetectorStatics+0)+6*A_PtrSize), "ptr", FaceDetectorStatics, "ptr*", FaceDetector)   ; CreateAsync
      WaitForAsync(FaceDetector)
      DllCall(NumGet(NumGet(FaceDetectorStatics+0)+7*A_PtrSize), "ptr", FaceDetectorStatics, "ptr*", ReadOnlyList)   ; GetSupportedBitmapPixelFormats
      DllCall(NumGet(NumGet(ReadOnlyList+0)+7*A_PtrSize), "ptr", ReadOnlyList, "int*", count)   ; count
      loop % count
      {
         DllCall(NumGet(NumGet(ReadOnlyList+0)+6*A_PtrSize), "ptr", ReadOnlyList, "int", A_Index-1, "uint*", BitmapPixelFormat)   ; get_Item
         SupportedBitmapPixelFormats .= "|" BitmapPixelFormat "|"
      }
      ObjRelease(FaceDetectorStatics)
      ObjRelease(ReadOnlyList)
   }
   if (SubStr(file, 2, 1) != ":")
      file := A_ScriptDir "\" file
   if !FileExist(file) or InStr(FileExist(file), "D")
   {
      msgbox File "%file%" does not exist
      ExitApp
   }   
   VarSetCapacity(GUID, 16)
   DllCall("ole32\CLSIDFromString", "wstr", IID_RandomAccessStream := "{905A0FE1-BC53-11DF-8C49-001E4FC686DA}", "ptr", &GUID)
   DllCall("ShCore\CreateRandomAccessStreamOnFile", "wstr", file, "uint", Read := 0, "ptr", &GUID, "ptr*", IRandomAccessStream)
   DllCall(NumGet(NumGet(BitmapDecoderStatics+0)+14*A_PtrSize), "ptr", BitmapDecoderStatics, "ptr", IRandomAccessStream, "ptr*", BitmapDecoder)   ; CreateAsync
   WaitForAsync(BitmapDecoder)
   BitmapFrame := ComObjQuery(BitmapDecoder, IBitmapFrame := "{72A49A1C-8081-438D-91BC-94ECFC8185C6}")
   DllCall(NumGet(NumGet(BitmapFrame+0)+12*A_PtrSize), "ptr", BitmapFrame, "uint*", width)   ; get_PixelWidth
   DllCall(NumGet(NumGet(BitmapFrame+0)+13*A_PtrSize), "ptr", BitmapFrame, "uint*", height)   ; get_PixelHeight
   DllCall(NumGet(NumGet(BitmapFrame+0)+8*A_PtrSize), "ptr", BitmapFrame, "uint*", BitmapPixelFormat)   ; get_BitmapPixelFormat
   BitmapFrameWithSoftwareBitmap := ComObjQuery(BitmapDecoder, IBitmapFrameWithSoftwareBitmap := "{FE287C9A-420C-4963-87AD-691436E08383}")
   if (height > maxheight)
   {
      DllCall(NumGet(NumGet(BitmapEncoderStatics+0)+15*A_PtrSize), "ptr", BitmapEncoderStatics, "ptr", IRandomAccessStream, "ptr", BitmapDecoder, "ptr*", BitmapEncoder)   ; CreateForTranscodingAsync
      WaitForAsync(BitmapEncoder)
      DllCall(NumGet(NumGet(BitmapEncoder+0)+15*A_PtrSize), "ptr", BitmapEncoder, "ptr*", BitmapTransform)   ; BitmapTransform
      DllCall(NumGet(NumGet(BitmapTransform+0)+7*A_PtrSize), "ptr", BitmapTransform, "int", floor(maxheight/height*width))   ; put_ScaledWidth
      DllCall(NumGet(NumGet(BitmapTransform+0)+9*A_PtrSize), "ptr", BitmapTransform, "int", maxheight)   ; put_ScaledHeight
      DllCall(NumGet(NumGet(BitmapFrameWithSoftwareBitmap+0)+8*A_PtrSize), "ptr", BitmapFrameWithSoftwareBitmap, "uint", BitmapPixelFormat, "uint", Premultiplied := 0, "ptr", BitmapTransform, "uint", IgnoreExifOrientation := 0, "uint", DoNotColorManage := 0, "ptr*", SoftwareBitmap)   ; GetSoftwareBitmapTransformedAsync
   }
   else
      DllCall(NumGet(NumGet(BitmapFrameWithSoftwareBitmap+0)+6*A_PtrSize), "ptr", BitmapFrameWithSoftwareBitmap, "ptr*", SoftwareBitmap)   ; GetSoftwareBitmapAsync
   WaitForAsync(SoftwareBitmap)
   if !InStr(SupportedBitmapPixelFormats, "|" BitmapPixelFormat "|")
   {
      DllCall(NumGet(NumGet(SoftwareBitmapStatics+0)+7*A_PtrSize), "ptr", SoftwareBitmapStatics, "ptr", SoftwareBitmap, "uint", Gray8 := 62, "ptr*", SoftwareBitmapTemp)   ; Convert
      Close := ComObjQuery(SoftwareBitmap, IClosable := "{30D5A829-7FA4-4026-83BB-D75BAE4EA99E}")
      DllCall(NumGet(NumGet(Close+0)+6*A_PtrSize), "ptr", Close)   ; Close
      ObjRelease(Close)
      ObjRelease(SoftwareBitmap)
      SoftwareBitmap := SoftwareBitmapTemp
   }
   DllCall(NumGet(NumGet(FaceDetector+0)+6*A_PtrSize), "ptr", FaceDetector, ptr, SoftwareBitmap, "ptr*", DetectedFaceList)   ; DetectFacesAsync
   WaitForAsync(DetectedFaceList)
   DllCall(NumGet(NumGet(DetectedFaceList+0)+7*A_PtrSize), "ptr", DetectedFaceList, "int*", count)   ; count
   loop % count
   {
      varsetcapacity(bounds, 16, 0)
      DllCall(NumGet(NumGet(DetectedFaceList+0)+6*A_PtrSize), "ptr", DetectedFaceList, "int", A_Index-1, "ptr*", DetectedFace)   ; get_Item
      DllCall(NumGet(NumGet(DetectedFace+0)+6*A_PtrSize), "ptr", DetectedFace, "ptr", &bounds)   ; BitmapBounds
      x := numget(bounds, 0, "uint")
      y := numget(bounds, 4, "uint")
      width := numget(bounds, 8, "uint")
      height := numget(bounds, 12, "uint")
      result .= "face" A_Index ": x=" x ", y=" y ", width=" width ", height=" height "`n"
      ObjRelease(DetectedFace)
   }
   Close := ComObjQuery(IRandomAccessStream, IClosable := "{30D5A829-7FA4-4026-83BB-D75BAE4EA99E}")
   DllCall(NumGet(NumGet(Close+0)+6*A_PtrSize), "ptr", Close)   ; Close
   ObjRelease(Close)
   Close := ComObjQuery(SoftwareBitmap, IClosable := "{30D5A829-7FA4-4026-83BB-D75BAE4EA99E}")
   DllCall(NumGet(NumGet(Close+0)+6*A_PtrSize), "ptr", Close)   ; Close
   ObjRelease(Close)
   ObjRelease(IRandomAccessStream)
   ObjRelease(BitmapDecoder)
   ObjRelease(BitmapFrame)
   if (height > maxheight)
   {
      ObjRelease(BitmapEncoder)
      ObjRelease(BitmapTransform)
   }
   ObjRelease(BitmapFrameWithSoftwareBitmap)
   ObjRelease(SoftwareBitmap)
   ObjRelease(DetectedFaceList)
   return result
}



CreateClass(string, interface, ByRef Class)
{
   CreateHString(string, hString)
   VarSetCapacity(GUID, 16)
   DllCall("ole32\CLSIDFromString", "wstr", interface, "ptr", &GUID)
   result := DllCall("Combase.dll\RoGetActivationFactory", "ptr", hString, "ptr", &GUID, "ptr*", Class, "uint")
   if (result != 0)
   {
      if (result = 0x80004002)
         msgbox No such interface supported
      else if (result = 0x80040154)
         msgbox Class not registered
      else
         msgbox error: %result%
      ExitApp
   }
   DeleteHString(hString)
}

CreateHString(string, ByRef hString)
{
    DllCall("Combase.dll\WindowsCreateString", "wstr", string, "uint", StrLen(string), "ptr*", hString)
}

DeleteHString(hString)
{
   DllCall("Combase.dll\WindowsDeleteString", "ptr", hString)
}

WaitForAsync(ByRef Object)
{
   AsyncInfo := ComObjQuery(Object, IAsyncInfo := "{00000036-0000-0000-C000-000000000046}")
   loop
   {
      DllCall(NumGet(NumGet(AsyncInfo+0)+7*A_PtrSize), "ptr", AsyncInfo, "uint*", status)   ; IAsyncInfo.Status
      if (status != 0)
      {
         if (status != 1)
         {
            DllCall(NumGet(NumGet(AsyncInfo+0)+8*A_PtrSize), "ptr", AsyncInfo, "uint*", ErrorCode)   ; IAsyncInfo.ErrorCode
            msgbox AsyncInfo status error: %ErrorCode%
            ExitApp
         }
         ObjRelease(AsyncInfo)
         break
      }
      sleep 10
   }
   DllCall(NumGet(NumGet(Object+0)+8*A_PtrSize), "ptr", Object, "ptr*", ObjectResult)   ; GetResults
   ObjRelease(Object)
   Object := ObjectResult
}