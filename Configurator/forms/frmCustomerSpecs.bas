Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    TabularFamily =21
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8640
    DatasheetFontHeight =10
    ItemSuffix =68
    Left =5235
    Top =5460
    Right =15660
    Bottom =11985
    DatasheetGridlinesColor =12632256
    Filter ="[CustomerPartNumber] like \"*\""
    OrderBy ="[tblCustomerSpecs].[CustomerPartNumber], [tblCustomerSpecs].[Customer]"
    RecSrcDt = Begin
        0x89dbdd2a54bbe440
    End
    RecordSource ="tblCustomerSpecs"
    Caption ="Item Search"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =2040
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =2220
                    Top =1680
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Text12"
                    Caption ="Document"
                    LayoutCachedLeft =2220
                    LayoutCachedTop =1680
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =1680
                    Width =1920
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label21"
                    Caption ="Customer Part #"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1680
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =1980
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =240
                    Top =60
                    Width =2940
                    Name ="Box35"
                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =780
                End
                Begin Label
                    OverlapFlags =85
                    Left =4200
                    Top =1680
                    Width =1620
                    Height =300
                    FontSize =10
                    FontWeight =500
                    Name ="Label52"
                    Caption ="Customer"
                    LayoutCachedLeft =4200
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =1980
                End
                Begin Line
                    OverlapFlags =85
                    Left =180
                    Top =1380
                    Width =7860
                    Name ="Line58"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1380
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =1380
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5520
                    Top =900
                    Width =1320
                    TabIndex =1
                    Name ="btnRemoveFilter"
                    Caption ="Remove Filter"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =900
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =1260
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =18
                    Left =360
                    Top =180
                    Width =2700
                    Height =495
                    FontSize =20
                    FontWeight =500
                    Name ="lblItemSearch"
                    Caption ="Customer Specs"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =360
                    LayoutCachedTop =180
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =675
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =4020
                    Top =900
                    Width =1320
                    TabIndex =2
                    Name ="btnFilter"
                    Caption ="Apply Filter"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =900
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =1260
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =180
                    Top =900
                    Width =3600
                    Height =300
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    Name ="ItemNoSearch"
                    Format ="@;\" <enter search text>\""
                    FontName ="Times New Roman"

                    LayoutCachedLeft =180
                    LayoutCachedTop =900
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =1200
                End
            End
        End
        Begin Section
            Height =426
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2220
                    Top =60
                    Width =1860
                    Height =300
                    FontSize =9
                    FontWeight =500
                    Name ="CusomerSpec"
                    ControlSource ="CustomerSpec"
                    AfterUpdate ="[Event Procedure]"
                    OnDirty ="[Event Procedure]"

                    LayoutCachedLeft =2220
                    LayoutCachedTop =60
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="CustomerPartNumber"
                    ControlSource ="CustomerPartNumber"

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6480
                    Width =486
                    Height =426
                    TabIndex =2
                    Name ="btnFind"
                    Caption ="Command13"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000001e0000001c0000000100180000000000100a0000c40e0000c40e0000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffefefefefefefefefefefefefefefefefefefefefefefe ,
                        0xfefefefefefefefefefefefefefefefefefefeffffffffffffffffffffffffff ,
                        0xffff0000fffffffffffffffffffffffffffffffffffffffffffdfdfd9a9a9aff ,
                        0xfffffdfdfdfefefefefefefefefefefefefefefefefefefefefefefefefefefe ,
                        0xfefefefefefefefefefefefefdfdfdfffffff3f3f3adadadf4f4f4ffffff0000 ,
                        0xfffffffffffffffffffffffffffffffffffffffffff9f9f9121212fcfcfcfefe ,
                        0xfef7f7f7f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f8f8f8f9f9f9ffffffe1e1e13d3d3de4e4e4ffffff0000ffffffff ,
                        0xfffffffffffffffffffffffffffffffffffbfbfb2b2b2bfdfdfdffffffd4d4d4 ,
                        0xd9d9d9dadadadadadadadadadadadadadadadadadadadadadadadadadadadada ,
                        0xdad7d7d7e2e2e2fefefee7e7e7515151e8e8e8ffffff0000ffffffffffffffff ,
                        0xfffffffffffffffffffffffffffafafb29292afdfdfdffffffcbccccd3d3d4d3 ,
                        0xd3d4d3d3d4d3d3d4d3d3d4d3d3d4d3d3d4d3d3d4d3d3d4d3d3d4d3d3d4d0d0d1 ,
                        0xdbdbdcfefefee7e7e7505050e8e8e8ffffff0000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffff363421ffffffffffffe3e0cee9e6d6e8e7d6e8e7 ,
                        0xd6e8e7d6e8e7d6e8e7d6e8e7d6e8e7d6e8e7d6e8e7d6e9e7d6e7e5d3eceadeff ,
                        0xfffee7e7e7505050e8e8e8ffffff0000ffffffecedf7939bd9828cd58891d788 ,
                        0x91d68791de6974ce0002588c96dc8a92d87178b9747cbe757cbe747cbe747cbe ,
                        0x747cbe747cbe747cbe747cbe747cbe747cbe757cbe6b73ba9298ccf5f5fce9e9 ,
                        0xe8505050e8e8e8ffffff0000ffffffd6d6f01e29b10007a50512a9040fad0008 ,
                        0x87acb0d69ba3dc0607a70411a90310a90311aa0412aa0413aa0412aa0411aa04 ,
                        0x11aa0411aa0412aa0412aa0310aa0311aa0003a44549bce7eaf9ebebe8505050 ,
                        0xe8e8e8ffffff0000ffffffd9d8f11c2ab10008a60006a70000ab000579dadbe1 ,
                        0xd4d9f100009c0000a20012af0012aa0004a60000a30008ad0001a80009a7000a ,
                        0xad0000a30008a60216ab0115aa0004a54347bce8ebfaebebe8505050e8e8e8ff ,
                        0xffff0000ffffffd8daf21d2cb90008ad404fc0535fd2000a7cd0d1ddecedfc60 ,
                        0x6ed46d7bd8000493010db1424fc27884d92b36a04e5aba2c3dc7212fa47481d6 ,
                        0x2e3fc0000cb00519b50008ae464ac3e9ebfbebebe8505050e8e8e8ffffff0000 ,
                        0xffffffd9daf11d2bc00008ad8f98d2caccf114197dc8c8dcffffffbbbac2ffff ,
                        0xff9da0c73743bcffffffbfbee19c9abefffff76c78e3676b93ffffff939be300 ,
                        0x00b0051ab90007b4454bc7e9ecfcebebe8505050e8e8e8ffffff0000ffffffd9 ,
                        0xd8f31e2fc2000cba1023bf0e22c9000684d5d6e0d0d7f8000094dee2f2fafafb ,
                        0x7478bbfbfaf53a3db16466befffff75d6ce34e5594ffffff7b85dd0000b3061b ,
                        0xbe0009b9454bcae9ecfbebebe8505050e8e8e8ffffff0000ffffffd9d9f41e2f ,
                        0xc30010bd0115bd000ec5000a85d9d9e0ecf0fe4746c3f3f6fad5d8ef616bc4ff ,
                        0xfff97d83c48e93c5fffffa7781cd9092b6ffffffb0b5e50514bb0115bf000aba ,
                        0x454ccce9ecfbebebe8505050e8e8e8ffffff0000ffffffd9daf31e2fc70010c1 ,
                        0x061cc30117ca010e97afb3ddd5d9f3d4dae9e0e4fb3c41ca1e2fc9bec4ead1d5 ,
                        0xe9dcdcd9ffffffa3a6c1c7c7d5ffffffffffff3f4dca000bbf000bbf454ccfe9 ,
                        0xecfcebebe8505050e8e8e8ffffff0000ffffffd9d9f31e30cb0011c5051cc703 ,
                        0x1ac80117bf2d3bce2a39cc2030c63946d30000bd0014c41628c6303dd35b65c6 ,
                        0xa2a6da515fd8242ea9a4a9e3adb1dd212fc20012c5000cc4464dd3e9ecfcebeb ,
                        0xe7505050e8e8e8ffffff0000ffffffdad8f31e31cd0013c7051dc9051cc9041c ,
                        0xcb0013c70012c70013c70010c5051eca041cca0015c80010c60311bf0f12b804 ,
                        0x12c3000dc40f18b11f2a9d0010be021ccb000ec6464fd5eaecfdebebe7505050 ,
                        0xe8e8e8ffffff0000ffffffd4d7f21022ca0000c4000ac6000ac6000ac6000cc7 ,
                        0x0017d4000ac7000ac6000dca000ecb000fcb000fcb000cca0006c8000ac9000f ,
                        0xcd0008c60008c3000cca000fca0000c73d43d5e6eafceaeae7515151e8e8e8ff ,
                        0xffff0000ffffffe5e4f86b78df5562da5c6adc5b69dc5d6bde5a68de0e1d9265 ,
                        0x72e75f6de04150c44652c94652c94652c94653c94653ca4653c94653c95361d8 ,
                        0x5462da4f5dd3505dd44552cf7c86e1fbfdfffffefc414141e3e3e3ffffff0000 ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffe36331efffffdffff ,
                        0xffd5d4bfdddac8dddbc7dddbc7dddbc7dddbc7dedbc8e6e4d19b98849e9b87b4 ,
                        0xb39fb0ae9bb0ac99afad9ebcbcb96e6e6ea5a5a5fdfdfdffffff0000ffffffff ,
                        0xfffffffffffffffffffffffffffffffffffbfbfb292a2afefefdffffffc0c0c0 ,
                        0xc8c8c9c8c8c9c8c8c9c8c8c9c8c8c9c9c9cad8d8d95253535454557777787070 ,
                        0x717171727575757b7b7b111111c6c6c6ffffffffffff0000ffffffffffffffff ,
                        0xfffffffffffffffffffffffffffbfbfb2a2a2afdfdfdffffffc7c7c7cccccccd ,
                        0xcdcdcdcdcdcdcdcdcdcdcdcececedadada656565ecececffffffffffffffffff ,
                        0xffffff737373bebebeffffffffffffffffff0000ffffffffffffffffffffffff ,
                        0xfffffffffffffffffffbfbfb2a2a2afdfdfdffffffcbcbcbd2d2d2d2d2d2d2d2 ,
                        0xd2d2d2d2d2d2d2d3d3d3dfdfdf646464d5d5d5fffffff5f5f5bfbfbfa3a3a3e1 ,
                        0xe1e1ffffffffffffffffffffffff0000ffffffffffffffffffffffffffffffff ,
                        0xfffffffffffbfbfb2a2a2afdfdfdffffffd0d0d0d5d5d5d5d5d5d5d5d5d5d5d5 ,
                        0xd5d5d5d6d6d6e2e2e2656565d6d6d6ffffffcecece414141f1f1f1ffffffffff ,
                        0xffffffffffffffffffff0000ffffffffffffffffffffffffffffffffffffffff ,
                        0xfffbfbfb2a2a2afdfdfdffffffd3d3d3d8d8d8d8d8d8d8d8d8d8d8d8d8d8d8d9 ,
                        0xd9d9e6e6e6696969e2e2e2fefefe787878a8a8a8ffffffffffffffffffffffff ,
                        0xffffffffffff0000fffffffffffffffffffffffffffffffffffffffffffbfbfb ,
                        0x212121fffffffffffff5f5f5f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f7ffff ,
                        0xff7f7f7f5c5c5cc0c0c0c2c2c2ffffffffffffffffffffffffffffffffffffff ,
                        0xffff0000fffffffffffffffffffffffffffffffffffffffffffdfdfd4d4d4df1 ,
                        0xf1f1ebebebececececececececececececececececececebebebf9f9f97a7a7a ,
                        0x818181f2f2f2ffffffffffffffffffffffffffffffffffffffffffffffff0000 ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffa0a0a0a2a2a2a1a1 ,
                        0xa1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a2a2a29f9f9fa9a9a9ffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffff0000ffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffff9f9f9f7f7f7f7f7f7f7f7f7 ,
                        0xf7f7f7f7f7f7f7f7f7f7f7f7f7f7f7f8f8f8f7f7f7fafafaffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff0000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    Picture ="pdf_file.bmp"

                    LayoutCachedLeft =6480
                    LayoutCachedWidth =6966
                    LayoutCachedHeight =426
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4200
                    Top =60
                    Width =2100
                    Height =300
                    FontSize =9
                    FontWeight =500
                    TabIndex =3
                    Name ="Customer"
                    ControlSource ="Customer"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =60
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8040
                    Width =366
                    Height =366
                    TabIndex =4
                    Name ="cmdDelete"
                    Caption ="Command12"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada177adada77da1dad1177adad17ad11da7117dad71ada ,
                        0x111da1177d117dad1111d7117711dada11111d11111dadad1111da71117adada ,
                        0x111d77111177adad11d711da71177ada1dadadada71177addadadadadad11ada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"

                    LayoutCachedLeft =8040
                    LayoutCachedWidth =8406
                    LayoutCachedHeight =366
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =1500
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =300
                    Top =660
                    Width =1320
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1920
                    Top =660
                    Width =1320
                    TabIndex =1
                    Name ="Refresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3600
                    Top =660
                    Width =1320
                    TabIndex =2
                    Name ="btnEdit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5220
                    Top =660
                    Width =1500
                    TabIndex =3
                    Name ="Command67"
                    Caption ="Add New Record"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5220
                    LayoutCachedTop =660
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =1020
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Database
Dim oldDocument As String

Private Sub btnClose_Click()
    DoCmd.Close
End Sub

Private Sub btnEdit_Click()
    EditField (True)

End Sub

Private Sub btnFilter_Click()

    Dim strFilter As String
    strFilter = "[CustomerPartNumber] like """ & Me.ItemNoSearch & "*"""
    strFilter = strFilter & " OR [CustomerSpec] like """ & Me.ItemNoSearch & "*"""
    strFilter = strFilter & " OR [Customer] like """ & Me.ItemNoSearch & "*"""

    Me.Filter = strFilter
    Me.FilterOn = True
    Call Refresh_Click

End Sub

Private Sub btnFind_Click()

    Dim Customer As String
    Dim CustomerPartNumber As String

    On Error GoTo ErrorHandler

    DoCmd.OpenForm "frmCustomerSpecList"

    Customer = Me.Customer
    CustomerPartNumber = Me.CustomerPartNumber

    Call ShowCustomerSpec(CustomerPartNumber, Customer)
    Forms.frmCustomerSpecList.bxSpecList.Selected(0) = True
    Forms.frmCustomerSpecList.Customer = Customer

ErrorExit:
    Exit Sub

ErrorHandler:
    DoCmd.Close acForm, "frmCustomerSpecList"
    MsgBox ("No customer spec found"), vbCritical
    Resume ErrorExit:


End Sub

Private Sub btnRemoveFilter_Click()
    Me.ItemNoSearch.value = Null
    Me.ItemNoSearch.Requery
    Me.Filter = ""
    Me.FilterOn = False
    Call Refresh_Click
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you SURE you want to delete this record?" & vbCrLf & _
              "This will permanently delete the record.", vbYesNo, "2nd Delete Confirmation") = vbYes Then
        DoCmd.SetWarnings False
        DoCmd.OpenQuery "qryDELETECustomerSpec"
        DoCmd.SetWarnings True
        Me.Requery
    End If
End Sub

Private Sub Command67_Click()
    DoCmd.OpenForm "frmCustomerSpecsNewRecord"
End Sub

Private Sub CusomerSpec_AfterUpdate()
    Call UpdateCustomerSpecName(Me.Customer, oldDocument, Me.CustomerSpec)

End Sub

Private Sub CusomerSpec_Dirty(Cancel As Integer)
    oldDocument = Me.CusomerSpec
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim strFilter As String
    strFilter = "[CustomerPartNumber] like """ & Me.ItemNoSearch & "*"""

    Me.Filter = strFilter
    Me.FilterOn = True
    Call Refresh_Click

    EditField (False)


End Sub

Private Sub Refresh_Click()
    Me.Requery
    Me.Refresh
End Sub

Private Sub EditField(lck As Boolean)


    Dim ctl As control
    For Each ctl In Me.Controls
        If ctl.ControlType = acTextBox Then
            If ctl.Locked <> lck Then
                ctl.Locked = lck
            End If
        End If
    Next ctl

End Sub
