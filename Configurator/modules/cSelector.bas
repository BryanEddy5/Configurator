Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'*****************************************************
'http://www.kleindata.ca/Tools.aspx?toolname=selector
'******************************************************

'events
Public Event Selected(value As String)

'constants
Private Const FORM_NAME = "fSelector"
'Private Const FORM_NAME = "frmSelector"
Private Const MIN_X = 3.5
Private Const MIN_Y = 4.5

'inputs
Private m_prompt As String
Private m_title As String
Private m_sql As String
Private m_cols As String
Private m_width As Long
Private m_height As Long
Private m_findID As Long
Private m_dbs As DAO.Database
Private m_formName$

'outputs
Private m_value As Variant
Private m_valueArray As Variant

'objects
'Private m_builder As cListViewBuilder
Private WithEvents m_frm As Form_fSelector
Attribute m_frm.VB_VarHelpID = -1
'Private WithEvents m_frm As Form_frmSelector


'********************************************************************
'*                                                                                                                                   Oct 7 2009
'*       Properties>
'*
'********************************************************************
Friend Property Get Prompt() As String
   Prompt = m_prompt
End Property

Friend Property Get Title() As String
   Title = m_title
End Property

Friend Property Get sql() As String
   sql = m_sql
End Property

Friend Property Get ColumnWidths() As String
   ColumnWidths = m_cols
End Property

Friend Property Get Width() As Long
   Width = m_width
End Property

Friend Property Get Height() As Long
   Height = m_height
End Property

Friend Property Get FindItemID() As Long
   FindItemID = m_findID
End Property

'********************************************************************
'*                                                                                                                                   Oct 7 2009
'*       Object Properties>
'*
'********************************************************************
Friend Property Get Form() As Form_fSelector
   Set Form = m_frm
End Property
Friend Property Set Form(frm As Form_fSelector)
   Set m_frm = frm
End Property

Friend Property Get Database() As DAO.Database
   If m_dbs Is Nothing Then Set m_dbs = CurrentDb
   Set Database = m_dbs
End Property

'Property Get ListViewBuilder() As cListViewBuilder
'   If m_builder Is Nothing Then Set m_builder = New cListViewBuilder
'   Set ListViewBuilder = m_builder
'End Property
   
'********************************************************************
'*                                                                                                                                   Oct 7 2009
'*       Return Values>
'*
'********************************************************************
Property Get Cancelled() As Boolean
   Cancelled = IsNull(m_value) Or IsEmpty(m_value)
End Property

Property Get ReturnValue() As Variant
   ReturnValue = m_value
End Property

Property Get ReturnValueArray() As Variant
   ReturnValueArray = m_valueArray
End Property
   
'********************************************************************
'*                                                                                                                                   Oct 7 2009
'*       Class Event Handlers
'*
'********************************************************************
Private Sub Class_Terminate()
   'make sure the form is closed when the object goes out of scope
   DoCmd.Close acForm, FORM_NAME, acSaveNo
End Sub


'********************************************************************
'*                                                                                                                                   Oct 7 2009
'*       Methods>
'*
'********************************************************************
Public Function Load( _
   Prompt As String, _
   Title As String, _
   sql As String, _
   Optional ColumnWidths As String, _
   Optional inchWidth As Single = 5.5, _
   Optional inchHeight As Single = 6.5, _
   Optional FindItemID As Long, _
   Optional dbs As DAO.Database _
   ) As cSelector
   
On Error GoTo handler
   'enforce form minimum sizes
   If inchWidth < MIN_X Then inchWidth = MIN_X
   If inchHeight < MIN_Y Then inchHeight = MIN_Y
   'save parameters
   m_prompt = Prompt
   m_title = Title
   m_sql = sql
   m_cols = ColumnWidths
   m_width = inchWidth * 1440
   m_height = inchHeight * 1440
   m_findID = FindItemID
   Set m_dbs = dbs
   
   Set Load = Me
   Exit Function
   
handler:
   Err.Raise Err, Err.Source & " cSelector.Load()"
End Function

Public Function Show() As Variant
On Error GoTo handler
   'A problem we need to solve is that the fSelector form opens as a dialog and we can't pass
   'it very rich information.  What we can do though, is put this object (Me) in a public place so
   'the Selector form can read it, so what we do is add Me to a public dictionary on the System form.
   LibSys.AddToDictionary "Selector", Me
   'Then we open fSelector as a dialog--so it is modal and popup
   DoCmd.OpenForm FORM_NAME, , , , , acDialog   'execution stops here until the form is hidden by the user
   'And we remove Me from the dictionary when the form loses focus
   LibSys.RemoveFromDictionary "Selector"
   
   'set return values
   m_value = Me.Form.value
   Show = Me.Form.value
   'set return value array
   m_valueArray = Me.Form.ValueArray

   'ensure form is closed and other objects are closed
   DoCmd.Close acForm, FORM_NAME
   
   'tidy up
   Set m_frm = Nothing
   'Set m_builder = Nothing
   Me.Database.Close
   Set m_dbs = Nothing
   
   Exit Function

handler:
   Show = Null
   m_value = Null
   DoCmd.Close acForm, FORM_NAME
   Err.Raise Err, Err.Source & " cSelector.Show()"
End Function

'********************************************************************
'*                                                                                                                                   Aug 20 2011
'*       Selector Form Events>
'*
'********************************************************************
Private Sub m_frm_Selected(value As String)
'  Handles and just re-raises the selected event
   RaiseEvent Selected(value)
End Sub