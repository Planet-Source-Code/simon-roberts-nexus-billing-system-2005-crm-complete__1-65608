VERSION 5.00
Object = "{79423413-BFDF-483D-BBB2-0D3B88187EB4}#1.0#0"; "whois.ocx"
Begin VB.Form frmWhois 
   BackColor       =   &H0066B0DD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WhoIs"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "frmWhois.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H0066B0DD&
      Height          =   5775
      Left            =   90
      TabIndex        =   4
      Top             =   2160
      Width           =   9075
      Begin VB.TextBox Text2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0043D143&
         Height          =   5415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   210
         Width           =   8805
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0066B0DD&
      Caption         =   "Domain You Are Searching for"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   90
      TabIndex        =   2
      Top             =   1140
      Width           =   9075
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Press Enter to search for domain on the server of your choice..."
         Top             =   300
         Width           =   8715
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0066B0DD&
      Caption         =   "Whois Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   9075
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   8775
      End
   End
   Begin WhoIsControl.WhoIs WhoIs1 
      Left            =   7050
      Top             =   1110
      _ExtentX        =   873
      _ExtentY        =   873
      Server          =   ""
      Query           =   ""
   End
End
Attribute VB_Name = "frmWhois"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmWhois"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of project alpha                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    
        
        
        Combo1.AddItem "whois.alabanza.com"
        Combo1.AddItem "whois.compuserve.com"
        Combo1.AddItem "whois.discount-domain.com"
        Combo1.AddItem "dns411.com"
        Combo1.AddItem "whois.domaindiscover.com"
        Combo1.AddItem "whois.domainpeople.com"
        Combo1.AddItem "whois.easyspace.com"
        Combo1.AddItem "whois.enom.com"
        Combo1.AddItem "whois.geektools.com"
        Combo1.ListIndex = Combo1.ListCount - 1
        Combo1.AddItem "whois.ibm.com"
        Combo1.AddItem "whois.internetnamesww.com"
        Combo1.AddItem "whois.names4ever.com"
        Combo1.AddItem "whois.namesecure.com"
        Combo1.AddItem "whois.networksolutions.com"
        Combo1.AddItem "whois.pacbell.com"
        Combo1.AddItem "whois.register.com"
        Combo1.AddItem "whois.registrars.com"
        Combo1.AddItem "whois.sunquest.com"
        Combo1.AddItem "whois.berkeley.edu"
        Combo1.AddItem "www.binghamton.edu"
        Combo1.AddItem "finger.caltech.edu"
        Combo1.AddItem "csufresno.edu"
        Combo1.AddItem "csuhayward.edu"
        Combo1.AddItem "csus.edu"
        Combo1.AddItem "whois.cwru.edu"
        Combo1.AddItem "cc.fsu.edu"
        Combo1.AddItem "directory.gatech.edu"
        Combo1.AddItem "gettysburg.edu"
        Combo1.AddItem "gmu.edu"
        Combo1.AddItem "whois.dfci.harvard.edu"
        Combo1.AddItem "hmc.edu"
        Combo1.AddItem "indiana.edu"
        Combo1.AddItem "nii.isi.edu"
        Combo1.AddItem "whois.isi.edu"
        Combo1.AddItem "whois.messiah.edu"
        Combo1.AddItem "whois.rsmas.miami.edu"
        Combo1.AddItem "mit.edu"
        Combo1.AddItem "directory.msstate.edu"
        Combo1.AddItem "vax2.winona.msus.edu"
        Combo1.AddItem "info.nau.edu"
        Combo1.AddItem "whois.ncsu.edu"
        Combo1.AddItem "nd.edu"
        Combo1.AddItem "earth.njit.edu"
        Combo1.AddItem "vm1.nodak.edu"
        Combo1.AddItem "austin.onu.edu"
        Combo1.AddItem "ph.orst.edu"
        Combo1.AddItem "osu.edu"
        Combo1.AddItem "whois.oxy.edu"
        Combo1.AddItem "info.psu.edu"
        Combo1.AddItem "whois.cc.rochester.edu"
        Combo1.AddItem "whitepages.rutgers.edu"
        Combo1.AddItem "whois.sdsu.edu"
        Combo1.AddItem "stanford.edu"
        Combo1.AddItem "camis.stanford.edu"
        Combo1.AddItem "stjohns.edu"
        Combo1.AddItem "sunysb.edu"
        Combo1.AddItem "whois.bcm.tmc.edu"
        Combo1.AddItem "whois.ubalt.edu"
        Combo1.AddItem "directory.ucdavis.edu"
        Combo1.AddItem "uchicago.edu"
        Combo1.AddItem "ucsd.edu"
        Combo1.AddItem "weber.ucsd.edu"
        Combo1.AddItem "cgl.ucsf.edu"
        Combo1.AddItem "whois.uh.edu"
        Combo1.AddItem "whois.umass.edu"
        Combo1.AddItem "lookup.umd.edu"
        Combo1.AddItem "umn.edu"
        Combo1.AddItem "ns.unl.edu"
        Combo1.AddItem "whois.upenn.edu"
        Combo1.AddItem "x500.utexas.edu"
        Combo1.AddItem "netlib2.cs.utk.edu"
        Combo1.AddItem "whois.virginia.edu"
        Combo1.AddItem "whois.wfu.edu"
        Combo1.AddItem "wisc.edu"
        Combo1.AddItem "wpi.wpi.edu"
        Combo1.AddItem "ibc.wustl.edu"
        Combo1.AddItem "vm1.hqadmin.doe.gov"
        Combo1.AddItem "wp.doe.gov"
        Combo1.AddItem "llnl.gov"
        Combo1.AddItem "x500.arc.nasa.gov"
        Combo1.AddItem "x500.gsfc.nasa.gov"
        Combo1.AddItem "whois.hq.nasa.gov"
        Combo1.AddItem "x500.ivv.nasa.gov"
        Combo1.AddItem "whois.jpl.nasa.gov"
        Combo1.AddItem "x500.jsc.nasa.gov"
        Combo1.AddItem "larc.nasa.gov"
        Combo1.AddItem "whois.larc.nasa.gov"
        Combo1.AddItem "x500.msfc.nasa.gov"
        Combo1.AddItem "x500.ssc.nasa.gov"
        Combo1.AddItem "x500.wstf.nasa.gov"
        Combo1.AddItem "x500.nasa.gov"
        Combo1.AddItem "wp.nersc.gov"
        Combo1.AddItem "whois.nic.gov"
        Combo1.AddItem "seda.sandia.gov"
        Combo1.AddItem "whois.nic.mil"
        Combo1.AddItem "whois.nrl.navy.mil"
        Combo1.AddItem "whois.6bone.net"
        Combo1.AddItem "whois.abuse.net"
        Combo1.AddItem "whois.aco.net"
        Combo1.AddItem "whois.apnic.net"
        Combo1.AddItem "whois.arin.net"
        Combo1.AddItem "whois.aunic.net"
        Combo1.AddItem "whois.awregistry.net"
        Combo1.AddItem "whois.cary.net"
        Combo1.AddItem "whois.corenic.net"
        Combo1.AddItem "whois.crsnic.net"
        Combo1.AddItem "whois.cw.net"
        Combo1.AddItem "wp.es.net"
        Combo1.AddItem "whois.hinet.net"
        Combo1.AddItem "ds.internic.net"
        Combo1.AddItem "whois.internic.net"
        Combo1.AddItem "whois.ja.net"
        Combo1.AddItem "whois.krnic.net"
        Combo1.AddItem "whois.lac.net"
        Combo1.AddItem "companies.mci.net"
        Combo1.AddItem "whois.nameit.net"
        Combo1.AddItem "whois.netnames.net"
        Combo1.AddItem "whois.nomination.net"
        Combo1.AddItem "whois.nsiregistry.net"
        Combo1.AddItem "whois.oleane.net"
        Combo1.AddItem "whois.opensrs.net"
        Combo1.AddItem "pcdc.net"
        Combo1.AddItem "whois.ra.net"
        Combo1.AddItem "whois.ripe.net"
        Combo1.AddItem "whois.ripn.net"
        Combo1.AddItem "whois.thnic.net"
        Combo1.AddItem "whois.twnic.net"
        Combo1.AddItem "whois.dhs.org"
        Combo1.AddItem "whois.morris.org"
        Combo1.AddItem "whois.nic.ac"
        Combo1.AddItem "whois.nic.am"
        Combo1.AddItem "whois.nic.as"
        Combo1.AddItem "wp.tuwien.ac.at"
        Combo1.AddItem "whois.risc.uni-linz.ac.at"
        Combo1.AddItem "whois.wu-wien.ac.at"
        Combo1.AddItem "archie.au"
        Combo1.AddItem "whois.connect.com.au"
        Combo1.AddItem "whois.adelaide.edu.au"
        Combo1.AddItem "whois.monash.edu.au"
        Combo1.AddItem "uwa.edu.au"
        Combo1.AddItem "sserve.cc.adfa.oz.au"
        Combo1.AddItem "whois.kuleuven.ac.be"
        Combo1.AddItem "whois.belnet.be"
        Combo1.AddItem "whois.registro.br"
        Combo1.AddItem "whois.camosun.bc.ca"
        Combo1.AddItem "whois.canet.ca"
        Combo1.AddItem "whois.cdnnet.ca"
        Combo1.AddItem "whois.queensu.ca"
        Combo1.AddItem "ac.nsac.ns.ca"
        Combo1.AddItem "whois.unb.ca"
        Combo1.AddItem "panda1.uottawa.ca"
        Combo1.AddItem "dvinci.usask.ca"
        Combo1.AddItem "whois.usask.ca"
        Combo1.AddItem "phys.uvic.ca"
        Combo1.AddItem "whois.uwo.ca"
        Combo1.AddItem "whois.nic.cc"
        Combo1.AddItem "whois.nic.ch"
        Combo1.AddItem "whois.nic.ck"
        Combo1.AddItem "whois.nic.cl"
        Combo1.AddItem "whois.cnnic.net.cn"
        Combo1.AddItem "whois.ci.ucr.ac.cr"
        Combo1.AddItem "whois.cuni.cz"
        Combo1.AddItem "whois.mff.cuni.cz"
        Combo1.AddItem "www.fce.vutbr.cz"
        Combo1.AddItem "gopher.fme.vutbr.cz"
        Combo1.AddItem "whois.fee.vutbr.cz"
        Combo1.AddItem "whois.vutbr.cz"
        Combo1.AddItem "whois.fh-koeln.de"
        Combo1.AddItem "whois.fzi.de"
        Combo1.AddItem "hermes.informatik.htw-zittau.de"
        Combo1.AddItem "whois.nic.de"
        Combo1.AddItem "whois.th-darmstadt.de"
        Combo1.AddItem "whois.tu-chemnitz.de"
        Combo1.AddItem "whois.uni-regensburg.de"
        Combo1.AddItem "whois.uni-c.dk"
        Combo1.AddItem "whois.ut.ee"
        Combo1.AddItem "whois.eunet.es"
        Combo1.AddItem "whois.dit.upm.es"
        Combo1.AddItem "cs.hut.fi"
        Combo1.AddItem "oulu.fi"
        Combo1.AddItem "vtt.fi"
        Combo1.AddItem "whois.nic.fr"
        Combo1.AddItem "whois.nordnet.fr"
        Combo1.AddItem "whois.univ-lille1.fr"
        Combo1.AddItem "whois.hknic.net.hk"
        Combo1.AddItem "whois.registry.hm"
        Combo1.AddItem "whois.iisc.ernet.in"
        Combo1.AddItem "whois.ncst.ernet.in"
        Combo1.AddItem "isgate.is"
        Combo1.AddItem "isgate3.isnet.is"
        Combo1.AddItem "pgebrehiwot.iat.cnr.it"
        Combo1.AddItem "dsa.nis.garr.it"
        Combo1.AddItem "whois.nic.it"
        Combo1.AddItem "whois.nic.mx"
        Combo1.AddItem "whois.aist-nara.ac.jp"
        Combo1.AddItem "whois-server.l.chiba-u.ac.jp"
        Combo1.AddItem "whois.hiroshima-u.ac.jp"
        Combo1.AddItem "gopher.educ.cc.keio.ac.jp"
        Combo1.AddItem "whois.cc.keio.ac.jp"
        Combo1.AddItem "whois.cc.uec.ac.jp"
        Combo1.AddItem "whois.yamanashi.ac.jp"
        Combo1.AddItem "whois.nic.ad.jp"
        Combo1.AddItem "www.orions.ad.jp"
        Combo1.AddItem "whois.domain.kg"
        Combo1.AddItem "sorak.kaist.ac.kr"
        Combo1.AddItem "whois.nic.or.kr"
        Combo1.AddItem "whois.domain.kz"
        Combo1.AddItem "whois.nic.li"
        Combo1.AddItem "whois.nic.lk"
        Combo1.AddItem "www.restena.lu"
        Combo1.AddItem "whois.nic.mm"
        Combo1.AddItem "www.nic.mx"
        Combo1.AddItem "condor.dgsca.unam.mx"
        Combo1.AddItem "domain-registry.nl"
        Combo1.AddItem "whois.norid.no"
        Combo1.AddItem "whois.nic.nu"
        Combo1.AddItem "whois.canterbury.ac.nz"
        Combo1.AddItem "directory.vuw.ac.nz"
        Combo1.AddItem "waikato.ac.nz"
        Combo1.AddItem "whois.patho.gen.nz"
        Combo1.AddItem "whois.domainz.net.nz"
        Combo1.AddItem "whois.rcp.net.pe"
        Combo1.AddItem "whois.icm.edu.pl"
        Combo1.AddItem "whois.elka.pw.edu.pl"
        Combo1.AddItem "whois.ia.pw.edu.pl"
        Combo1.AddItem "whois.dns.pt"
        Combo1.AddItem "dsa.fccn.pt"
        Combo1.AddItem "chalmers.se"
        Combo1.AddItem "kth.se"
        Combo1.AddItem "whois.nic-se.se"
        Combo1.AddItem "sics.se"
        Combo1.AddItem "whois.nic.net.sg"
        Combo1.AddItem "whois.nic.sh"
        Combo1.AddItem "whois.uakom.sk"
        Combo1.AddItem "whois.nic.st"
        Combo1.AddItem "whois.adamsnames.tc"
        Combo1.AddItem "whois.nic.tj"
        Combo1.AddItem "whois.tonic.to"
        Combo1.AddItem "whois.metu.edu.tr"
        Combo1.AddItem "whois.seed.net.tw"
        Combo1.AddItem "whois.iii.org.tw"
        Combo1.AddItem "src.doc.ic.ac.uk"
        Combo1.AddItem "whois.lut.ac.uk"
        Combo1.AddItem "whois.nic.uk"
        Combo1.AddItem "dsa.shu.ac.uk"
        Combo1.AddItem "whois.state.ct.us"
        Combo1.AddItem "info.cnri.reston.va.us"
        Combo1.AddItem "whois.frd.ac.za"
        Combo1.AddItem "whois.und.ac.za"
        Combo1.AddItem "whois.co.za"

Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Text1_KeyPress"
    Const ContainerName = "frmWhois"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of project alpha                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    Select Case KeyAscii
    Case 13
        Screen.MousePointer = 11
        WhoIs1.Server = Combo1.Text
        WhoIs1.Query = Text1.Text
        WhoIs1.Connect
                
        Text2 = "Whois Query Started " & Format(sysNOW, "dddd dd-mm-yyyy ttttt")
        Text2 = Text2 + vbCrLf + "Query = " & Text1.Text
        Text2 = Text2 + vbCrLf
        Text2 = Text2 + vbCrLf + IIf(WhoIs1.Result = "", "No Result Returned Try a different Server", WhoIs1.Result)
        Screen.MousePointer = 0
   End Select
   
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub
