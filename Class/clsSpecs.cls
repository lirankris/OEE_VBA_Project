VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'class Module: Main_Line_spec
Option Explicit

Private Site_location As String
Private Location As String
Private line_name As String
Private Incharge As String
Private Pace As Integer
Private Box_pace As Integer
Private Pallet_pace As Integer
Private Size As Integer
Private I_sb As String
Private delta As Date
Private Box_Delta As Date
Private Pallet_Delta As Date
Private b_size As Integer
Private m_wieght As Integer

'/////////////////////////////////////////////M_site_location////////////////////////////////////////////////////////
' M_site_location = Main site location
Property Get M_site_location() As String
    M_site_location = Site_location
End Property

Property Let M_site_location(msl As String)
    Site_location = msl
End Property
'////////////////////////////////////////////////D_Location/////////////////////////////////////////////////////////
' D_Location = Division Location
Property Get D_Location() As String
    D_Location = Location
End Property

Property Let D_Location(dloc As String)
     Location = dloc
End Property
'//////////////////////////////////////////////D_line_name///////////////////////////////////////////////////////////
' D_line_name = Division line name
Property Get D_line_name() As String
    D_line_name = line_name
End Property

Property Let D_line_name(ln As String)
     line_name = ln
End Property
'//////////////////////////////////////////////D_Incharge////////////////////////////////////////////////////////////
' D_Incharge = Division Incharge
Property Get D_Incharge() As String
    D_Incharge = Incharge
End Property

Property Let D_Incharge(inc As String)
     Incharge = inc
End Property
'///////////////////////////////////////////////////////////////////////////Pace////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////C_pace//////////////////////////////////////////////////////////////
' C_pace = can pace
Property Get C_pace() As Integer
    C_pace = Pace
End Property

Property Let C_pace(pa As Integer)
     Pace = pa
End Property
'/////////////////////////////////////////////////B_pace/////////////////////////////////////////////////////////////
' B_pace = bag pace
Property Get B_pace() As Integer
    B_pace = Pace
End Property

Property Let B_pace(pa As Integer)
     Pace = pa
End Property
'////////////////////////////////////////////////P_box_pace//////////////////////////////////////////////////////////
' B_pace = bag pace
Property Get P_box_pace() As Integer
    P_box_pace = Box_pace
End Property

Property Let P_box_pace(bpa As Integer)
     Box_pace = bpa
End Property
'//////////////////////////////////////////////P_pallte_pace/////////////////////////////////////////////////////////
' B_pace = bag pace
Property Get P_pallte_pace() As Integer
    P_pallte_pace = Pallet_pace
End Property

Property Let P_pallte_pace(ppa As Integer)
     Pallet_pace = ppa
End Property
'///////////////////////////////////////////////////////////////////////////Substance///////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////C_sb//////////////////////////////////////////////////////////////
' C_sb = Can Substance Type
Property Get C_sb() As String
    C_sb = I_sb
End Property

Property Let C_sb(sb As String)
     I_sb = sb
End Property
'////////////////////////////////////////////////////B_sb///////////////////////////////////////////////////////////
' B_sb = Bag Substance Type
Property Get B_sb() As String
    B_sb = I_sb
End Property

Property Let B_sb(sb As String)
     I_sb = sb
End Property
'///////////////////////////////////////////////////////////////////////////Size////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////C_p_size//////////////////////////////////////////////////////////
' C_p_size = can pack size
Property Get C_p_size() As Integer
    C_p_size = Size
End Property

Property Let C_p_size(ps As Integer)
     Size = ps
End Property

'////////////////////////////////////////////////////B_p_size////////////////////////////////////////////////////////
' B_p_size = bag pack size
Property Get B_p_size() As Integer
    B_p_size = b_size
End Property

Property Let B_p_size(bs As Integer)
     b_size = bs
End Property

'///////////////////////////////////////////////////////////////////////////max_w////////////////////////////////////////////////////////////////////////////////////////////
' max_w = max pack weight
Property Get max_w() As Integer
    max_w = m_wieght
End Property

Property Let max_w(mx As Integer)
     m_wieght = mx
End Property


