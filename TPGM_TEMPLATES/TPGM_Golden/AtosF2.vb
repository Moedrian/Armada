Option Strict Off
Option Explicit On
Module modAtosF2
'ifndef r3cuChannelsH
Public Const r3cuChannelsH As Integer = 0
'ifdef __cplusplus
Declare Function TpDisconnectAllAbus Lib "r3cu_COMMON_200.DLL" ( ByVal RowMask As Integer ) As Integer 
Declare Function TpConnectAbus Lib "r3cu_COMMON_200.DLL" ( ByRef TpArray As Integer , ByVal RowMask As Integer ) As Integer 
Declare Function TpDisconnectAbus Lib "r3cu_COMMON_200.DLL" ( ByRef TpArray As Integer , ByVal RowMask As Integer ) As Integer 
Declare Function TpStrConnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal TpList As String , ByVal RowMask As Integer ) As Integer 
Declare Function TpStrDisconnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal TpList As String , ByVal RowMask As Integer ) As Integer 
Declare Function TpConnectUabus Lib "r3cu_COMMON_200.DLL" ( ByRef TpArray As Integer , ByVal RowMask As Integer ) As Integer 
Declare Function TpDisconnectUabus Lib "r3cu_COMMON_200.DLL" ( ByRef TpArray As Integer , ByVal RowMask As Integer ) As Integer 
Declare Function TpStrConnectUabus Lib "r3cu_COMMON_200.DLL" ( ByVal TpList As String , ByVal RowMask As Integer ) As Integer 
Declare Function TpStrDisconnectUabus Lib "r3cu_COMMON_200.DLL" ( ByVal TpList As String , ByVal RowMask As Integer ) As Integer 
Declare Function ChanDisconnectAllAbus Lib "r3cu_COMMON_200.DLL" ( ByVal RowMask As Integer ) As Integer 
Declare Function ChanConnectAbus Lib "r3cu_COMMON_200.DLL" ( ByRef ChanArray As Short , ByVal RowMask As Integer ) As Integer 
Declare Function ChanDisconnectAbus Lib "r3cu_COMMON_200.DLL" ( ByRef ChanArray As Short , ByVal RowMask As Integer ) As Integer 
Declare Function ChanStrConnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal ChanList As String , ByVal RowMask As Integer ) As Integer 
Declare Function ChanStrDisconnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal ChanList As String , ByVal RowMask As Integer ) As Integer 
Declare Function ChanConnectUabus Lib "r3cu_COMMON_200.DLL" ( ByRef ChanArray As Short , ByVal RowMask As Integer ) As Integer 
Declare Function ChanDisconnectUabus Lib "r3cu_COMMON_200.DLL" ( ByRef ChanArray As Short , ByVal RowMask As Integer ) As Integer 
Declare Function ChanStrConnectUabus Lib "r3cu_COMMON_200.DLL" ( ByVal ChanList As String , ByVal RowMask As Integer ) As Integer 
Declare Function ChanStrDisconnectUabus Lib "r3cu_COMMON_200.DLL" ( ByVal ChanList As String , ByVal RowMask As Integer ) As Integer 
Declare Function ChAbusModeSet Lib "r3cu_COMMON_200.DLL" ( ByVal Mode As Integer ) As Integer 
Declare Function ScnExtInstrConnect Lib "r3cu_COMMON_200.DLL" ( ByVal InstrId As Integer , ByVal vbPoint As Integer ) As Integer 
Declare Function ScnExtInstrDisconnect Lib "r3cu_COMMON_200.DLL" ( ByVal InstrId As Integer , ByVal vbPoint As Integer ) As Integer 
Declare Function ScnUabusModeSet Lib "r3cu_COMMON_200.DLL" ( ByVal InstrId As Integer , ByVal Mode As Integer ) As Integer 
Declare Function ScnConnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal InstrId As Integer , ByVal Row As Integer ) As Integer 
Declare Function ScnDisconnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal InstrId As Integer , ByVal Row As Integer ) As Integer 
Declare Function InsulatedChConnect Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function InsulatedChDisconnect Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Sub ChConnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal ChanStr As String , ByVal Row As Integer )   
Declare Sub ChDisconnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal ChanStr As String , ByVal Row As Integer )   
Declare Sub TestPointConnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal TpList As String , ByVal RowMask As Integer )   
Declare Sub TestPointDisconnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal TpList As String , ByVal RowMask As Integer )   
Declare Sub TestPointDisconnectAllAbus Lib "r3cu_COMMON_200.DLL" ( ByVal RowMask As Integer )   
'if !defined(__COMMON_H)
Public Const uuCOMMON_H As Integer = 0
'include "r3cuCommonConsts.h"
'include "r3cuCommonTypes.h"
'include "r3cuTestInfo.h"
'include "r3cuIct.h"
'include "r3cuUtility.h"
'include "r3cuChannels.h"
'include "r3cuFlyingProbe.h"
'ifdef __cplusplus
Declare Function CommonLibraryAttach Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function CommonLibraryDetach Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Sub CommonDefaultsSet Lib "r3cu_COMMON_200.DLL" ( ByVal vbEvent As Integer )   
Declare Function fProbeAssignCreate Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
'if !defined(__COMMON_CONSTS_H)
Public Const uuCOMMON_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_COMMON_UNIT As String =            "RP3_CommonUnit"
'endif
'ifndef r3cuIctH
Public Const r3cuIctH As Integer = 0
'ifndef __cplusplus
    'pragma pack (1)
'endif
''typedef 'struct
''{
'  long TprjId;           // Test project identifier
'  long TplanId;          // Testplan identifier
'  long Site;             // Site number
'  char TaskName[32];     // Task name
'  long TaskNumber;       // Task number
'  long TestNumber;       // Test number in the task
'  long UniqueTestId;     // Unique test identifier
'  char TestRemark[128];  // Test remark
'  long TestResult;       // Test result
'  double MeasuredValue;  // Measured value
'  double HighThreshold;  // High threshold
'  double LowThreshold;   // Low threshold
'  char Unit[8];          // Measure unit ("ohm", "V", …)
''} TTestMeas;
'ifndef __cplusplus
    'pragma pack ()
'endif
'ifdef __cplusplus
Declare Function RunAnlTest Lib "r3cu_COMMON_200.DLL" ( ByVal TestPlanName As String , ByVal StartTest As Integer , ByVal EndTest As Integer ) As Integer 
Declare Function RunAnlTaskNum Lib "r3cu_COMMON_200.DLL" ( ByVal TestplanName As String , ByVal StartTask As Integer , ByVal EndTask As Integer ) As Integer 
Declare Function RunAnlTask Lib "r3cu_COMMON_200.DLL" ( ByVal TestplanName As String , ByVal StartTaskName As String , ByVal EndTaskName As String ) As Integer 
Declare Function RunAnlTaskLabel Lib "r3cu_COMMON_200.DLL" ( ByVal TestplanName As String , ByVal StartTaskLabel As String , ByVal EndTaskLabel As String ) As Integer 
Declare Function RunAnlConcurrentSet Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function RunAnlWaitEndOfTplan Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function RunAIct Lib "r3cu_COMMON_200.DLL" ( ByVal TestplanName As String , ByVal StartTest As Integer , ByVal EndTest As Integer ) As Integer 
Declare Function RunATask Lib "r3cu_COMMON_200.DLL" ( ByVal TestplanName As String , ByVal StartTask As Integer , ByVal EndTask As Integer ) As Integer 
Declare Sub AnlTaskMeasStoreEnable Lib "r3cu_COMMON_200.DLL" ( )   
Declare Sub AnlTaskMeasStoreDisable Lib "r3cu_COMMON_200.DLL" ( )   
'  long WINAPI ReadAnlTaskMeas (char *TestplanName, char *StartTaskName, char *EndTaskName, TTestMeas *TestMeasArray, long ArraySize, long *StoredMeasSize);
'  long WINAPI ReadAnlTaskLabelMeas (char *TestplanName, char *StartTaskLabel, char *EndTaskLabel, TTestMeas *TestMeasArray, long ArraySize, long *StoredMeasSize);
Declare Function SkipFirstMovement Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
'ifndef r3cuTestInfoH
Public Const r3cuTestInfoH As Integer = 0
'ifdef __cplusplus
Declare Function TpgmNameRead Lib "r3cu_COMMON_200.DLL" ( ByVal TpgmName As String ) As Integer 
Declare Function TpgmFolderRead Lib "r3cu_COMMON_200.DLL" ( ByVal TpgmFolder As String ) As Integer 
Declare Function TpgmDirRead Lib "r3cu_COMMON_200.DLL" ( ByVal TpgmFolder As String ) As Integer 
Declare Function TpgmPathRead Lib "r3cu_COMMON_200.DLL" ( ByVal TpgmPath As String ) As Integer 
Declare Function TpgmPauseSet Lib "r3cu_COMMON_200.DLL" ( ByVal pPauseParameter As String ) As Integer 
Declare Function TpgmStartTimeRead Lib "r3cu_COMMON_200.DLL" ( ByRef pTm_hour As Integer , ByRef pTm_min As Integer , ByRef pTm_sec As Integer ) As Integer 
Declare Function TpgmStartDateRead Lib "r3cu_COMMON_200.DLL" ( ByRef pTm_mday As Integer , ByRef pTm_mon As Integer , ByRef pTm_year As Integer ) As Integer 
Declare Function CheckPointSet Lib "r3cu_COMMON_200.DLL" ( ByVal Site As Integer , ByVal CheckPointId As Integer ) As Integer 
Declare Function OperatorRead Lib "r3cu_COMMON_200.DLL" ( ByVal OperatorName As String ) As Integer 
Declare Function SerialNumberRead Lib "r3cu_COMMON_200.DLL" ( ByVal Site As Integer , ByVal SerialNumber As String ) As Integer 
Declare Function SerialNumberWrite Lib "r3cu_COMMON_200.DLL" ( ByVal Site As Integer , ByVal SerialNumber As String ) As Integer 
Declare Function LotNumberRead Lib "r3cu_COMMON_200.DLL" ( ByVal LotNumber As String ) As Integer 
Declare Function LotRead Lib "r3cu_COMMON_200.DLL" ( ByVal LotNumber As String ) As Integer 
Declare Function ProductRead Lib "r3cu_COMMON_200.DLL" ( ByVal ProductName As String ) As Integer 
Declare Function LotNumberWrite Lib "r3cu_COMMON_200.DLL" ( ByVal LotNumber As String ) As Integer 
Declare Function LotWrite Lib "r3cu_COMMON_200.DLL" ( ByVal LotNumber As String ) As Integer 
Declare Function ProductWrite Lib "r3cu_COMMON_200.DLL" ( ByVal ProductName As String ) As Integer 
Declare Function HoldRegisterRead Lib "r3cu_COMMON_200.DLL" ( ByVal RegNum As Integer , ByRef Value As Double , ByVal Udm As String ) As Integer 
Declare Function ActiveHeadRead Lib "r3cu_COMMON_200.DLL" ( ByRef Head As Integer ) As Integer 
Declare Function HeadBusyTimeRead Lib "r3cu_COMMON_200.DLL" ( ByVal pHead As Integer , ByRef pBusyTime As Integer ) As Integer 
Declare Function TesterNumberRead Lib "r3cu_COMMON_200.DLL" ( ByRef TesterNum As Integer ) As Integer 
Declare Function SystemSerialNumberRead Lib "r3cu_COMMON_200.DLL" ( ByVal SysSn As String ) As Integer 
Declare Function HoldRegisterWrite Lib "r3cu_COMMON_200.DLL" ( ByVal RegNum As Integer , ByVal Value As Double , ByVal Udm As String ) As Integer 
Declare Function UseSiteRead Lib "r3cu_COMMON_200.DLL" ( ByRef Site As Integer ) As Integer 
Declare Function BoardInTestRead Lib "r3cu_COMMON_200.DLL" ( ByRef Site As Integer ) As Integer 
Declare Function UseSiteWrite Lib "r3cu_COMMON_200.DLL" ( ByVal Site As Integer ) As Integer 
Declare Function BoardInTestWrite Lib "r3cu_COMMON_200.DLL" ( ByVal Site As Integer ) As Integer 
Declare Function UseProjectWrite Lib "r3cu_COMMON_200.DLL" ( ByVal Tprj As Integer ) As Integer 
Declare Function FrameCodeRead Lib "r3cu_COMMON_200.DLL" ( ByVal FrameCode As String ) As Integer 
Declare Function SitesEnabledGet Lib "r3cu_COMMON_200.DLL" ( ByVal pSitesEnabled As String ) As Integer 
Declare Function VariantCompositionRead Lib "r3cu_COMMON_200.DLL" ( ByVal VariantComposition As String ) As Integer 
Declare Function VariantCompositionWrite Lib "r3cu_COMMON_200.DLL" ( ByVal VariantComposition As String ) As Integer 
Declare Function VariantCompositionNumRead Lib "r3cu_COMMON_200.DLL" ( ByRef VariantCompositionNumber As Integer ) As Integer 
Declare Function VariantCheck Lib "r3cu_COMMON_200.DLL" ( ByVal VariantNum As Integer ) As Integer 
Declare Function BoardCodeRead Lib "r3cu_COMMON_200.DLL" ( ByRef BoardCode As Integer ) As Integer 
Declare Function FixtureCodeRead Lib "r3cu_COMMON_200.DLL" ( ByRef FixtureCode As Integer ) As Integer 
Declare Function FixtureMaintenanceRead Lib "r3cu_COMMON_200.DLL" ( ByRef FixtureMaintenance As Integer ) As Integer 
Declare Function ContrastPlateCodeRead Lib "r3cu_COMMON_200.DLL" ( ByRef ContrastPlateCode As Integer ) As Integer 
Declare Function BoardCodeBayRead Lib "r3cu_COMMON_200.DLL" ( ByVal BayNr As Integer , ByRef BoardCode As Integer ) As Integer 
Declare Function FixtureCodeBayRead Lib "r3cu_COMMON_200.DLL" ( ByVal BayNr As Integer , ByRef FixtureCode As Integer ) As Integer 
Declare Function FixtureMaintenanceBayRead Lib "r3cu_COMMON_200.DLL" ( ByVal BayNr As Integer , ByRef FixtureMaintenance As Integer ) As Integer 
Declare Function ContrastPlateCodeBayRead Lib "r3cu_COMMON_200.DLL" ( ByVal BayNr As Integer , ByRef ContrastPlateCode As Integer ) As Integer 
Declare Function AnlTempFailTest Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function AnlFailTest Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function FuncTempFailTest Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function FuncFailTest Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function SiteResultRead Lib "r3cu_COMMON_200.DLL" ( ByVal Site As Integer ) As Integer 
Declare Function SiteResultReadV2 Lib "r3cu_COMMON_200.DLL" ( ByVal TprjPos As Integer , ByVal Site As Integer ) As Integer 
Declare Function SiteResultWrite Lib "r3cu_COMMON_200.DLL" ( ByVal Site As Integer , ByVal SiteResult As Integer ) As Integer 
Declare Function SiteResultWriteV2 Lib "r3cu_COMMON_200.DLL" ( ByVal TprjPos As Integer , ByVal Site As Integer , ByVal SiteResult As Integer ) As Integer 
Declare Function SerialNumberResultRead Lib "r3cu_COMMON_200.DLL" ( ByVal SnPos As Integer ) As Integer 
Declare Function OrganizerResultSet Lib "r3cu_COMMON_200.DLL" ( ByVal OrganizerResult As Integer ) As Integer 
Declare Function OrganizerResultSetV2 Lib "r3cu_COMMON_200.DLL" ( ByVal TprjPos As Integer , ByVal OrganizerResult As Integer ) As Integer 
Declare Sub ObsDatalogTest Lib "r3cu_COMMON_200.DLL" ( ByVal TestNumber As Integer , ByVal DrawingReference As String , ByVal Remark As String , ByVal TestResult As Integer , ByVal MeasuredValueStr As String , ByVal ThrLowStr As String , ByVal ThrHighStr As String , ByVal TpListStr As String , ByVal SiteNum As Integer )   
Declare Sub ObsDatalogTestV2 Lib "r3cu_COMMON_200.DLL" ( ByVal TprjPos As Integer , ByVal TestNumber As Integer , ByVal DrawingReference As String , ByVal Remark As String , ByVal TestResult As Integer , ByVal MeasuredValueStr As String , ByVal ThrLowStr As String , ByVal ThrHighStr As String , ByVal TpListStr As String , ByVal SiteNum As Integer )   
Declare Sub FuncDatalogTest Lib "r3cu_COMMON_200.DLL" ( ByVal TaskNumber As Integer , ByVal TestNumberInTask As Integer , ByVal UniqueTestId As Integer , ByVal TestNumber As Integer , ByVal DrawingReference As String , ByVal Remark As String , ByVal TestResult As Integer , ByVal MeasuredValueStr As String , ByVal ThrLowStr As String , ByVal ThrHighStr As String , ByVal TpListStr As String , ByVal SiteNum As Integer )   
Declare Sub BscanDatalogTest Lib "r3cu_COMMON_200.DLL" ( ByVal TaskNumber As Integer , ByVal TestNumberInTask As Integer , ByVal UniqueTestId As Integer , ByVal TestNumber As Integer , ByVal DrawingReference As String , ByVal Remark As String , ByVal TestResult As Integer , ByVal MeasuredValueStr As String , ByVal ThrLowStr As String , ByVal ThrHighStr As String , ByVal TpListStr As String , ByVal SiteNum As Integer )   
Declare Sub ObpDatalogTest Lib "r3cu_COMMON_200.DLL" ( ByVal TaskNumber As Integer , ByVal TestNumberInTask As Integer , ByVal UniqueTestId As Integer , ByVal TestNumber As Integer , ByVal DrawingReference As String , ByVal Remark As String , ByVal TestResult As Integer , ByVal MeasuredValueStr As String , ByVal ThrLowStr As String , ByVal ThrHighStr As String , ByVal TpListStr As String , ByVal SiteNum As Integer )   
Declare Function UseBayRead Lib "r3cu_COMMON_200.DLL" ( ByRef BayNo As Integer ) As Integer 
Declare Function SyncParallelExec Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function TpgmMultiprjRead Lib "r3cu_COMMON_200.DLL" ( ByRef pIsMultiprj As Integer ) As Integer 
Declare Function UseSiteBayRead Lib "r3cu_COMMON_200.DLL" ( ByVal pBay As Integer , ByRef pSite As Integer ) As Integer 
Declare Function BoardInTestBayRead Lib "r3cu_COMMON_200.DLL" ( ByVal pBay As Integer , ByRef pSite As Integer ) As Integer 
Declare Function TpgmSerialNumberListRead Lib "r3cu_COMMON_200.DLL" ( ByVal pSnList As String ) As Integer 
Declare Function ExtConnectUabusBmu Lib "r3cu_COMMON_200.DLL" ( ByVal ExtInstrId As Integer , ByVal RowPoint1 As Integer , ByVal RowPoint2 As Integer , ByVal RowPoint3 As Integer , ByVal RowPoint4 As Integer ) As Integer 
Declare Function ExtDisconnectUabusBmu Lib "r3cu_COMMON_200.DLL" ( ByVal ExtInstrId As Integer , ByVal ExtInstrPoint As Integer ) As Integer 
Declare Function GndConnectUabusBmu Lib "r3cu_COMMON_200.DLL" ( ByVal Row As Integer ) As Integer 
Declare Function GndDisconnectUabusBmu Lib "r3cu_COMMON_200.DLL" ( ByVal Row As Integer ) As Integer 
Declare Function BmuMove Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer , ByVal BmuPadId As Integer ) As Integer 
Declare Function BmuMoveUpDown Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer , ByVal BmuUpDown As Integer ) As Integer 
Declare Function BmuMoveXY Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer , ByVal CoordX As Integer , ByVal CoordY As Integer ) As Integer 
Declare Function BmuConnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal Row As Integer , ByVal Abus As Integer ) As Integer 
Declare Function BmuDisconnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal Row As Integer , ByVal Abus As Integer ) As Integer 
Declare Function BmuTpConnectUabus Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer , ByVal Testpoint As Integer , ByVal Row As Integer ) As Integer 
Declare Function BmuTpDisconnectUabus Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer , ByVal Testpoint As Integer , ByVal Row As Integer ) As Integer 
Declare Function BmuTpDisconnectAllUabus Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer , ByVal Testpoint As Integer ) As Integer 
Declare Function BmuPinConnectUabus Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer , ByVal Pin As Integer , ByVal Row As Integer ) As Integer 
Declare Function BmuPinDisconnectUabus Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer , ByVal Pin As Integer , ByVal Row As Integer ) As Integer 
Declare Function BmuPinDisconnectAllUabus Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer , ByVal Pin As Integer ) As Integer 
Declare Function BmuDisconnectAllUabus Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer ) As Integer 
Declare Function BmuDiagChanConnectUabus Lib "r3cu_COMMON_200.DLL" ( ByRef ChanArray As Short , ByVal Row As Integer ) As Integer 
Declare Function BmuDiagChanDisconnectUabus Lib "r3cu_COMMON_200.DLL" ( ByRef ChanArray As Short , ByVal Row As Integer ) As Integer 
Declare Function BmuDiagChanDisconnectAll Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function BmuPeConnect Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer , ByVal BitMask As Integer ) As Integer 
Declare Function BmuPeDisconnect Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer , ByVal BitMask As Integer ) As Integer 
Declare Function BmuGndConnect Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer ) As Integer 
Declare Function BmuGndDisconnect Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer ) As Integer 
Declare Function BmuEthernetConnect Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer ) As Integer 
Declare Function BmuEthernetDisconnect Lib "r3cu_COMMON_200.DLL" ( ByVal BmuId As Integer ) As Integer 
Declare Function FingerMoveXY Lib "r3cu_COMMON_200.DLL" ( ByVal FingerId As Integer , ByVal CoordX As Integer , ByVal CoordY As Integer ) As Integer 
Declare Function SemaphoreSet Lib "r3cu_COMMON_200.DLL" ( ByVal pLightMask As Integer , ByVal pBuzzerMask As Integer ) As Integer 
'ifndef r3cuUtilityH
Public Const r3cuUtilityH As Integer = 0
'ifdef __cplusplus
Declare Sub IdleMessage Lib "r3cu_COMMON_200.DLL" ( )   
'  char *WINAPI A2S (WORD *ItemArray, char *ConvString, long OptimizeFlag);
'  char *WINAPI A2STrace (long *ItemArray, char *ConvString, long OptimizeFlag);
'  char *WINAPI WA2STrace (WORD *ItemArray, char *ConvString, long OptimizeFlag);
Declare Function S2A Lib "r3cu_COMMON_200.DLL" ( ByVal strList As String , ByRef List As Short , ByVal MaxListSize As Integer ) As Integer 
'  'ifdef __cplusplus
'    long WINAPI S2WA (char *strList, WORD *List, long MaxListSize, WORD MaxItemValue=&HFFFF);
Declare Function S2WA Lib "r3cu_COMMON_200.DLL" ( ByVal strList As String , ByRef List As Short , ByVal MaxListSize As Integer , ByVal MaxItemValue As Short ) As Integer 
Declare Function S2DWA Lib "r3cu_COMMON_200.DLL" ( ByVal strList As String , ByRef List As Integer , ByVal MaxListSize As Integer ) As Integer 
'  char *WINAPI DWA2S (DWORD *ItemArray, char *ConvString, long OptimizeFlag);
Declare Function S2LA Lib "r3cu_COMMON_200.DLL" ( ByVal strList As String , ByRef List As Integer , ByVal MaxListSize As Integer ) As Integer 
'  char *WINAPI LA2S (long *ItemArray, char *ConvString, long OptimizeFlag);
'  char *WINAPI Mt2S (int MeasType, char *StrMeasure);
'  'ifdef __cplusplus
'    char *WINAPI Mv2S (double MeasValue, int MeasType, char *StrMeasure, int NumOfDecimal=3);
'    char *WINAPI Mv2S (double MeasValue, int MeasType, char *StrMeasure, int NumOfDecimal);
'  char *TrcGetRowMask (long RowMask, char *TrcString);
'  char *TrcGetRowId (long RowId, char *TrcString);
'  char *TrcGetAbusId (long AbusId, char *TrcString);
'  char *TrcGetBmu (long BmuId, char *TrcString);
'ifndef r3cuFlyingProbeH
Public Const r3cuFlyingProbeH As Integer = 0
Public Const PRB_PARK As Integer =        0
Public Const PRB_NOT_USED As Integer =   -1
Public Const PRB_UP As Integer =         -2
Public Const PRB_DOWN As Integer =       -3
Public Const PRB_NO_MOVE As Integer =    -4
Public Const PRB_UP_RETOUCH As Integer = -5
Public Const PRBMASK_PROBE1 As Integer = &H01
Public Const PRBMASK_PROBE2 As Integer = &H02
Public Const PRBMASK_PROBE3 As Integer = &H04
Public Const PRBMASK_PROBE4 As Integer = &H08
Public Const PRBMASK_PROBE5 As Integer = &H10
Public Const PRBMASK_PROBE6 As Integer = &H20
Public Const PRB_TOP As Integer =    0
Public Const PRB_BOTTOM As Integer = 1
'ifdef __cplusplus
Declare Function ProbeMove Lib "r3cu_COMMON_200.DLL" ( ByVal pTp1 As Integer , ByVal pTp2 As Integer , ByVal pTp3 As Integer , ByVal pTp4 As Integer , ByVal pTp5 As Integer , ByVal pTp6 As Integer , ByVal pTp7 As Integer , ByVal pTp8 As Integer ) As Integer 
Declare Function ProbeMoveXY Lib "r3cu_COMMON_200.DLL" ( ByVal pFirstMove As Integer , ByVal pProbesXYData As String , ByRef pTpList As Integer , ByVal pProbesMask As Integer ) As Integer 
Declare Sub ProbeTableLoad Lib "r3cu_COMMON_200.DLL" ( ByVal pTablePath As String , ByVal pAbsPathFlag As Integer )   
Declare Function ProbeTableExecute Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Sub ProbeTableStop Lib "r3cu_COMMON_200.DLL" ( )   
Declare Sub ProbeTableBoardOffsetSet Lib "r3cu_COMMON_200.DLL" ( ByVal pXoffset As Integer , ByVal pYoffset As Integer , ByVal pCosAlfa As Double , ByVal pSinAlfa As Double , ByVal pCoefAllX As Double , ByVal pCoefAllY As Double , ByVal pMultiFiducialZoneId As Integer )   
Declare Function ProbeMoveV2 Lib "r3cu_COMMON_200.DLL" ( ByVal pTp1 As Integer , ByVal pTp2 As Integer , ByVal pTp3 As Integer , ByVal pTp4 As Integer , ByVal pTp5 As Integer , ByVal pTp6 As Integer , ByVal pTp7 As Integer , ByVal pTp8 As Integer ) As Integer 
Declare Function ProbeAutomaticMoving Lib "r3cu_COMMON_200.DLL" ( ByRef pTpList As Integer , ByRef pAssignedProbeList As Integer ) As Integer 
Declare Function ProbeDigitalDSSet Lib "r3cu_COMMON_200.DLL" ( ByVal pVh As Double , ByVal pVhh As Double , ByVal pVth As Double ) As Integer 
Declare Function ProbeDigitalDSReset Lib "r3cu_COMMON_200.DLL" ( ByVal pProbe As Integer ) As Integer 
Declare Function ProbeDigitalDSClear Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function ProbeDigitalDriverSet Lib "r3cu_COMMON_200.DLL" ( ByVal pProbe As Integer , ByVal pLogicLevel As Integer , ByVal pOutImp As Double ) As Integer 
Declare Function ProbeDigitalSensorRead Lib "r3cu_COMMON_200.DLL" ( ByVal pProbe As Integer , ByRef pLevel As Integer ) As Integer 
Declare Function TestPointAttributeRead Lib "r3cu_COMMON_200.DLL" ( ByVal pTpNumber As Integer , ByRef pTpType As Integer , ByRef pChannel As Integer , ByRef pPrbAcc As Integer , ByRef pXcoord As Double , ByRef pYcoord As Double ) As Integer 
Declare Function ProbeRetouchShift Lib "r3cu_COMMON_200.DLL" ( ByVal pProbeId As Integer , ByVal pOffsetX As Integer , ByVal pOffsetY As Integer ) As Integer 
Declare Function ProbeRetouchZ Lib "r3cu_COMMON_200.DLL" ( ByVal pProbesMask As Integer ) As Integer 
Declare Function ProbeRetouchXY Lib "r3cu_COMMON_200.DLL" ( ByVal pProbesMask As Integer ) As Integer 
Declare Function Barcode2DRead Lib "r3cu_COMMON_200.DLL" ( ByVal pSide As Integer , ByVal pSpotLightOn As Integer , ByVal pDiffuseLightOn As Integer , ByVal pBrightness As Integer , ByVal pContrast As Integer , ByVal pXpos As Integer , ByVal pYpos As Integer , ByVal pDelay As Integer , ByVal pCodeRead As String ) As Integer 
Declare Function Barcode2DReadV2 Lib "r3cu_COMMON_200.DLL" ( ByVal pSide As Integer , ByVal pDirectLightOn As Integer , ByVal pServiceLightOn As Integer , ByVal pTangentialLightOn As Integer , ByVal pBrightness As Integer , ByVal pContrast As Integer , ByVal pXpos As Integer , ByVal pYpos As Integer , ByVal pDelay As Integer , ByVal pCodeRead As String ) As Integer 
Declare Function DiagProbeI2TRead Lib "r3cu_COMMON_200.DLL" ( ByVal pMotorId As Integer , ByRef pI2tValue As Double ) As Integer 
Declare Function DiagDoorOpen Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function DiagDoorClose Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function DiagProbeMove Lib "r3cu_COMMON_200.DLL" ( ByVal pOffsetMask As Integer , ByVal pProbeId As Integer , ByVal pX As Integer , ByVal pY As Integer , ByVal pZ As Integer , ByVal pZoffset As Integer , ByVal pMultiFiducialZoneId As Integer ) As Integer 
Declare Function DiagProbePositionReadV2 Lib "r3cu_COMMON_200.DLL" ( ByVal pProbeId As Integer , ByVal pOffsetMask As Integer , ByRef pX As Integer , ByRef pY As Integer , ByRef pZ As Integer ) As Integer 
Declare Function DiagProbePositionRead Lib "r3cu_COMMON_200.DLL" ( ByVal pProbeId As Integer , ByVal pReadMode As Integer , ByRef pX As Integer , ByRef pY As Integer , ByRef pZ As Integer ) As Integer 
Declare Function DiagAxisMoveDelta Lib "r3cu_COMMON_200.DLL" ( ByVal pProbeId As Integer , ByVal pAxisId As Integer , ByVal pDelta As Integer ) As Integer 
Declare Function DiagEscanLock Lib "r3cu_COMMON_200.DLL" ( ByVal pProbeId As Integer , ByVal pLockFlag As Integer ) As Integer 
Declare Function DiagEscanLockWithoutProbeMoving Lib "r3cu_COMMON_200.DLL" ( ByVal pLockFlag As Integer , ByVal pFrontRear As Integer ) As Integer 
Declare Function DiagBoardMarkerMove Lib "r3cu_COMMON_200.DLL" ( ByVal pMarkerId As Integer , ByVal pX As Integer , ByVal pY As Integer , ByVal pMultiFiducialZoneId As Integer ) As Integer 
Declare Function DiagBoardMarkerActivate Lib "r3cu_COMMON_200.DLL" ( ByVal pMarkerTime As Integer ) As Integer 
Declare Function DiagWarpageMove Lib "r3cu_COMMON_200.DLL" ( ByVal pWarpageId As Integer , ByVal pX As Integer , ByVal pY As Integer , ByVal pMultiFiducialZoneId As Integer ) As Integer 
Declare Function DiagWarpageQuoteRead Lib "r3cu_COMMON_200.DLL" ( ByVal pWarpageId As Integer , ByRef pQuote As Double ) As Integer 
Declare Function DiagConveyorBoardIn Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function DiagConveyorBoardOut Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function DiagConveyorWidthSet Lib "r3cu_COMMON_200.DLL" ( ByVal pWidth As Integer ) As Integer 
Declare Function DiagConveyorWidthSetV2 Lib "r3cu_COMMON_200.DLL" ( ByVal pWidth As Integer , ByRef pMarkOffsetX As Integer , ByRef pMarkOffsetY As Integer ) As Integer 
Declare Function DiagModeSet Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function DiagModeReset Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function DiagZAxisCurrentLimitSet Lib "r3cu_COMMON_200.DLL" ( ByVal pProbeId As Integer , ByVal pCurrentLimit As Double ) As Integer 
Declare Function DiagZAxisSpeedSet Lib "r3cu_COMMON_200.DLL" ( ByVal pProbeId As Integer , ByVal pSpeed As Double ) As Integer 
Declare Function DiagZAxisAccelerationSet Lib "r3cu_COMMON_200.DLL" ( ByVal pProbeId As Integer , ByVal pAcceleration As Double ) As Integer 
Declare Function DiagIsEscanLocked Lib "r3cu_COMMON_200.DLL" ( ByVal pProbeId As Integer , ByRef pLocked As Integer ) As Integer 
Declare Function ProbeConnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal pProbe As Integer , ByVal pAbusLine As Integer ) As Integer 
Declare Function ProbeDisconnectAbus Lib "r3cu_COMMON_200.DLL" ( ByVal pProbe As Integer , ByVal pAbusLine As Integer ) As Integer 
Declare Function ProbeClearAbus Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function NetTestPointListRead Lib "r3cu_COMMON_200.DLL" ( ByVal pNetName As String , ByRef pTpList As Integer , ByVal pTpListSize As Integer , ByRef pTpQtyRead As Integer ) As Integer 
Declare Function DiagProfileTimeRead Lib "r3cu_COMMON_200.DLL" ( ByVal pTimeType As Integer , ByRef pReadValue As Double ) As Integer 
Declare Function DiagProfileOn Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function DiagProfileOff Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function DiagAxisCamSet Lib "r3cu_COMMON_200.DLL" ( ByVal pProbeId As Integer , ByVal pAxisId As Integer , ByVal pValue As Double ) As Integer 
Declare Function DiagSetupLoad Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function ProbeMoveAtTestNumber Lib "r3cu_COMMON_200.DLL" ( ByVal pTestNumber As Integer ) As Integer 
Declare Function fProbeAssignDestroy Lib "r3cu_COMMON_200.DLL" ( ) As Integer 
Declare Function DiagFingerDown Lib "r3cu_COMMON_200.DLL" ( ByVal pProbeId As Integer ) As Integer 
Declare Function DiagFingerUp Lib "r3cu_COMMON_200.DLL" ( ByVal pProbeId As Integer ) As Integer 
Declare Function DiagClampLock Lib "r3cu_COMMON_200.DLL" ( ByVal pFront As Integer , ByVal pRear As Integer ) As Integer 
Declare Function DiagClampUnlock Lib "r3cu_COMMON_200.DLL" ( ByVal pFront As Integer , ByVal pRear As Integer ) As Integer 
'if !defined(__DCALOAD_H)
Public Const uuDCALOAD_H As Integer = 0
'include "r3cuDcALoadConsts.h"
'include "r3cuDcALoadTypes.h"
'ifdef __cplusplus
Declare Function DcALoadLibraryAttach Lib "r3cu_DCALOAD_200.DLL" ( ) As Integer 
Declare Function DcALoadLibraryDetach Lib "r3cu_DCALOAD_200.DLL" ( ) As Integer 
Declare Sub DcALoadDefaultsSet Lib "r3cu_DCALOAD_200.DLL" ( ByVal vbEvent As Integer )   
Declare Function DcALoadConnect Lib "r3cu_DCALOAD_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DcALoadDisconnect Lib "r3cu_DCALOAD_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DcALoadEnable Lib "r3cu_DCALOAD_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DcALoadDisable Lib "r3cu_DCALOAD_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DcALoadSet Lib "r3cu_DCALOAD_200.DLL" ( ByVal InstrId As Integer , ByVal VValue As Double , ByVal IValue As Double ) As Integer 
Declare Function DcALoadClear Lib "r3cu_DCALOAD_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DcALoadRippleMeasSet Lib "r3cu_DCALOAD_200.DLL" ( ByVal InstrId As Integer , ByVal Mode As Integer , ByVal Range As Integer ) As Integer 
Declare Function ReadDcALoadDvm Lib "r3cu_DCALOAD_200.DLL" ( ByVal InstrId As Integer , ByVal vbType As Integer , ByRef Value As Double ) As Integer 
Declare Function DcALoadSetV2 Lib "r3cu_DCALOAD_200.DLL" ( ByVal InstrId As Integer , ByVal VValue As Double , ByVal IValue As Double , ByVal OutFormat As Integer , ByVal SlewRate As Integer ) As Integer 
Declare Function DcALoadTimingSet Lib "r3cu_DCALOAD_200.DLL" ( ByVal InstrId As Integer , ByVal Timing As Integer ) As Integer 
Declare Function DcALoadReadbackSet Lib "r3cu_DCALOAD_200.DLL" ( ByVal InstrId As Integer , ByVal MeasMode As Integer , ByVal MeasType As Integer , ByVal Filter As Integer , ByVal Range As Integer ) As Integer 
'if !defined(__DCALOAD_CONSTS_H)
Public Const uuDCALOAD_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_DCALOAD_UNIT As String =            "RP3_DcALoadUnit"
'endif'if !defined(__DIG_H)
Public Const uuDIG_H As Integer = 0
'include "r3cuDigConsts.h"
'include "r3cuDigTypes.h"
'ifdef __cplusplus
Declare Function DigLibraryAttach Lib "r3cu_DIG_200.DLL" ( ) As Integer 
Declare Function DigLibraryDetach Lib "r3cu_DIG_200.DLL" ( ) As Integer 
Declare Sub DigDefaultsSet Lib "r3cu_DIG_200.DLL" ( ByVal vbEvent As Integer )   
Declare Function DigPatternLoad Lib "r3cu_DIG_200.DLL" ( ByVal Pattern As String , ByRef TestPoints As Integer ) As Integer 
Declare Function DigStrPatternLoad Lib "r3cu_DIG_200.DLL" ( ByVal Pattern As String , ByVal TestPoints As String ) As Integer 
Declare Function RunD Lib "r3cu_DIG_200.DLL" ( ByVal Mode As Integer , ByVal Timeout As Integer , ByRef Result As Integer , ByRef FailedPins As Integer , ByRef FailedPinsCnt As Integer , ByRef FailedStep As Integer ) As Integer 
Declare Function DigTpStuckSet Lib "r3cu_DIG_200.DLL" ( ByVal Level As Integer , ByRef TestPoints As Integer ) As Integer 
Declare Function DigStrTpStuckSet Lib "r3cu_DIG_200.DLL" ( ByVal Level As Integer , ByVal TestPoints As String ) As Integer 
Declare Function DigTpDriverSensorOn Lib "r3cu_DIG_200.DLL" ( ByVal Vh As Double , ByVal Vl As Double , ByVal Vref As Double , ByRef TestPoints As Integer ) As Integer 
Declare Function DigStrTpDriverSensorOn Lib "r3cu_DIG_200.DLL" ( ByVal Vh As Double , ByVal Vl As Double , ByVal Vref As Double , ByVal TestPoints As String ) As Integer 
Declare Function DigTpDriverSensorOff Lib "r3cu_DIG_200.DLL" ( ) As Integer 
Declare Function DigTpOutputImpedanceSet Lib "r3cu_DIG_200.DLL" ( ByVal Impedance As Integer , ByRef TestPoints As Integer ) As Integer 
Declare Function DigStrTpOutputImpedanceSet Lib "r3cu_DIG_200.DLL" ( ByVal Impedance As Integer , ByVal TestPoints As String ) As Integer 
Declare Sub ATOSC2_LibraryAttach Lib "r3cu_DIG_200.DLL" ( )   
Declare Sub VECTFLY_LibraryAttach Lib "r3cu_DIG_200.DLL" ( )   
Declare Sub VECTFLY_LibraryDetach Lib "r3cu_DIG_200.DLL" ( )   
Declare Function ReadDigPinPatternData Lib "r3cu_DIG_200.DLL" ( ByVal PatternName As String , ByVal PinNum As Short , ByVal DieNum As Short , ByVal FromStep As Integer , ByVal ToStep As Integer , ByVal RdData As String ) As Integer 
'  int WINAPI pDigPatternWr (char *PattName, WORD *PinList, char *StepFrom, char *StepTo, char *CharSeq, WORD *DieList);
'  int WINAPI pDigStepTimeSet (int SetupNum, double TestTime);
'  int WINAPI pDigEdgeSet (int SetupNum, int EdgeNum, double EdgeTime);
Declare Function ReadP40DigTiming Lib "r3cu_DIG_200.DLL" ( ) As Integer 
Declare Function DigTimAcquisitionSet Lib "r3cu_DIG_200.DLL" ( ByVal MainTp As Integer , ByVal StartSlope As Integer , ByVal AuxTp As Integer , ByVal EnableLevel As Integer , ByVal EnableTp As Integer , ByVal MeasNumber As Integer , ByVal Resolution As Integer ) As Integer 
Declare Function DigTimAcquisitionRead Lib "r3cu_DIG_200.DLL" ( ByVal Tp As Integer , ByVal Timeout As Integer , ByVal MeasArraySize As Integer , ByRef MeasCount As Integer , ByRef MeasLevel As Integer , ByRef MeasTime As Double ) As Integer 
Declare Function PluLinConnect Lib "r3cu_DIG_200.DLL" ( ByVal PortId As Integer , ByVal ChanTx As Short , ByVal ChanRx As Short ) As Integer 
Declare Function PluLinDisconnect Lib "r3cu_DIG_200.DLL" ( ByVal PortId As Integer ) As Integer 
Declare Function PluLinDisconnectAll Lib "r3cu_DIG_200.DLL" ( ) As Integer 
Declare Function PluLinAttributeSet Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer , ByVal AttributeId As Integer , ByVal Value As Integer ) As Integer 
Declare Function PluLinOpen Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer ) As Integer 
Declare Function PluLinClose Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer ) As Integer 
Declare Function PluLinFrameSend Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer , ByVal FrameId As Integer , ByVal FrameDataLen As Integer , ByVal FrameData As String , ByRef ResultsList As Integer ) As Integer 
Declare Function PluLinFrameSendReceive Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer , ByVal FrameId As Integer , ByVal FrameReadDataLen As Integer , ByRef ResultsList As Integer ) As Integer 
Declare Function PluLinFrameRead Lib "r3cu_DIG_200.DLL" ( ByVal PortId As Integer , ByRef FrameId As Integer , ByRef FrameDataLen As Integer , ByVal FrameData As String ) As Integer 
Declare Function PluSerialPortChAssign Lib "r3cu_DIG_200.DLL" ( ByVal PortId As Integer , ByVal ChanTx As Short , ByVal ChanRx As Short ) As Integer 
Declare Function PluSerialPortOpen Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer , ByVal BaudRate As Integer , ByVal DataBits As Integer , ByVal Parity As Integer , ByVal StopBits As Integer ) As Integer 
Declare Function PluSerialPortClose Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer ) As Integer 
Declare Function PluSerialPortSend Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer , ByVal vbData As String , ByVal DataLen As Integer , ByRef ResultsList As Integer ) As Integer 
Declare Function PluSerialPortReceive Lib "r3cu_DIG_200.DLL" ( ByVal PortId As Integer , ByVal vbData As String , ByVal MaxDataLen As Integer , ByVal EndChar As Integer , ByRef ReadDataLen As Integer , ByVal Timeout As Integer ) As Integer 
Declare Function PluSerComAsyncConfig Lib "r3cu_DIG_200.DLL" ( ByVal PortId As Integer , ByVal Baudrate As Integer , ByVal IdleLevel As Integer , ByVal ChanTx As Short , ByVal ChanRx As Short , ByVal BitSequenceData As Integer , ByVal BitSequenceQty As Integer ) As Integer 
Declare Function PluSerComOpen Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer ) As Integer 
Declare Function PluSerComClose Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer ) As Integer 
Declare Function PluSerComDataSend Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer , ByVal vbData As String , ByVal BitQty As Integer ) As Integer 
Declare Function PluSerComDataReceive Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer , ByVal BitQty As Integer ) As Integer 
Declare Function PluSerComDataCompare Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer , ByVal vbData As String , ByVal BitQty As Integer ) As Integer 
Declare Function PluSerComRun Lib "r3cu_DIG_200.DLL" ( ByRef PortIdList As Integer , ByVal RxStartConditionData As Integer , ByVal RxStartConditionBitQty As Integer , ByVal Timeout As Integer , ByRef Result As Integer ) As Integer 
Declare Function PluSerComRead Lib "r3cu_DIG_200.DLL" ( ByVal PortId As Integer , ByVal BitDataLen As Integer , ByVal vbData As String , ByRef ReadBitDataLen As Integer ) As Integer 
'if !defined(__DIG_CONSTS_H)
Public Const uuDIG_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_DIG_UNIT As String =            "RP3_DigUnit"
Public Const PT_HF_FREQREF As Double =    2.4576
Public Const PT_HF_MINFREQ As Integer =    10
Public Const PT_HF_MAXFREQ As Integer =    50
Public Const PT_HF_MINFVCO As Integer =    40
Public Const PT_HF_MAXFVCO As Integer =    120
Public Const PT_HF_MINP As Integer =       4
Public Const PT_HF_MAXP As Integer =       130
Public Const PT_HF_MINQ As Integer =       3
Public Const PT_HF_MAXQ As Integer =       129
Public Const PW_TT_ENABLE As Integer =             &H02           ' 0000 0010
Public Const PW_IO_ENABLE As Integer =             &H08           ' 0000 1000
Public Const PW_CDC_ENABLE As Integer =            &H20           ' 0010 0000
Public Const PW_PUD_ENABLE As Integer =            &H80           ' 1000 0000
Public Const PW_TT_HIGH As Integer =               &H01           ' 0000 0001
Public Const PW_IO_OUT As Integer =                &H04           ' 0000 0100
Public Const PW_CDC_CARE As Integer =              &H10           ' 0001 0000
Public Const PW_PUD_UP As Integer =                &H40           ' 0100 0000
'endif
'if !defined(__SAU_H)
Public Const uuSAU_H As Integer = 0
'include "r3cuSauConsts.h"
'include "r3cuSauTypes.h"
'ifdef __cplusplus
Declare Function SauLibraryAttach Lib "r3cu_SAU_200.DLL" ( ) As Integer 
Declare Function SauLibraryDetach Lib "r3cu_SAU_200.DLL" ( ) As Integer 
Declare Sub SauDefaultsSet Lib "r3cu_SAU_200.DLL" ( ByVal vbEvent As Integer )   
Declare Function BstConnectInterf Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer , ByVal vbPoint As Integer ) As Integer 
Declare Function BstDisconnectInterf Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer , ByVal vbPoint As Integer ) As Integer 
Declare Function BstConnectAbus Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer , ByVal HotRow As Integer , ByVal ColdRow As Integer ) As Integer 
Declare Function BstDisconnectAbus Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer , ByVal HotRow As Integer , ByVal ColdRow As Integer ) As Integer 
Declare Function BstReadbackVConnectMbus Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer , ByVal MbusLine As Integer ) As Integer 
Declare Function BstReadbackVDisconnectMbus Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function BstReadbackIConnectMbus Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer , ByVal MbusLine As Integer ) As Integer 
Declare Function BstReadbackIDisconnectMbus Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function BstExtInputConnect Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer , ByVal Source As Integer ) As Integer 
Declare Function BstExtInputDisconnect Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function BstDisconnectAll Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function BstSourceSet Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer , ByVal VValue As Double , ByVal IValue As Double , ByVal IRange As Integer , ByVal OutForm As Integer ) As Integer 
Declare Function BstTimingSet Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer , ByVal Flag As Integer ) As Integer 
Declare Function BstEnable Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function BstDisable Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function BstClear Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function BstSyncSet Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer , ByVal Flag As Integer , ByVal SbusLine As Integer ) As Integer 
Declare Function BstConnectUabusBmu Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer , ByVal RowHotP As Integer , ByVal RowHotPK As Integer , ByVal RowColdP As Integer , ByVal RowColdPK As Integer , ByVal RowHotS As Integer , ByVal RowColdS As Integer ) As Integer 
Declare Function BstDisconnectUabusBmu Lib "r3cu_SAU_200.DLL" ( ByVal InstrId As Integer , ByVal vbPoint As Integer ) As Integer 
Declare Sub BstVClear Lib "r3cu_SAU_200.DLL" ( ByVal BoosterV As Integer )   
Declare Sub BstVEnable Lib "r3cu_SAU_200.DLL" ( ByVal BoosterV As Integer )   
Declare Sub BstVDisable Lib "r3cu_SAU_200.DLL" ( ByVal BoosterV As Integer )   
Declare Sub BstVConnectAbus Lib "r3cu_SAU_200.DLL" ( ByVal BoosterV As Integer , ByVal vbPoint As Integer , ByVal Row As Integer )   
Declare Sub BstVDisconnectAbus Lib "r3cu_SAU_200.DLL" ( ByVal BoosterV As Integer , ByVal vbPoint As Integer , ByVal Row As Integer )   
Declare Sub BstVConnectInterf Lib "r3cu_SAU_200.DLL" ( ByVal BoosterV As Integer , ByVal vbPoint As Integer )   
Declare Sub BstVDisconnectInterf Lib "r3cu_SAU_200.DLL" ( ByVal BoosterV As Integer , ByVal vbPoint As Integer )   
Declare Sub BstVDisconnectAll Lib "r3cu_SAU_200.DLL" ( ByVal BoosterV As Integer )   
Declare Sub BstVSourceSet Lib "r3cu_SAU_200.DLL" ( ByVal BoosterV As Integer , ByVal Voltage As Double , ByVal Current As Double , ByVal OutFormat As Integer , ByVal VoltRef As Integer )   
Declare Sub BstVSyncSet Lib "r3cu_SAU_200.DLL" ( ByVal BoosterV As Integer , ByVal SyncRef As Integer )   
Declare Sub BstIClear Lib "r3cu_SAU_200.DLL" ( ByVal BoosterI As Integer )   
Declare Sub BstIEnable Lib "r3cu_SAU_200.DLL" ( ByVal BoosterI As Integer )   
Declare Sub BstIDisable Lib "r3cu_SAU_200.DLL" ( ByVal BoosterI As Integer )   
Declare Sub BstIConnectAbus Lib "r3cu_SAU_200.DLL" ( ByVal BoosterI As Integer , ByVal vbPoint As Integer , ByVal Row As Integer )   
Declare Sub BstIDisconnectAbus Lib "r3cu_SAU_200.DLL" ( ByVal BoosterI As Integer , ByVal vbPoint As Integer , ByVal Row As Integer )   
Declare Sub BstIConnectInterf Lib "r3cu_SAU_200.DLL" ( ByVal BoosterI As Integer , ByVal vbPoint As Integer )   
Declare Sub BstIDisconnectInterf Lib "r3cu_SAU_200.DLL" ( ByVal BoosterI As Integer , ByVal vbPoint As Integer )   
Declare Sub BstIDisconnectAll Lib "r3cu_SAU_200.DLL" ( ByVal BoosterI As Integer )   
Declare Sub BstISourceSet Lib "r3cu_SAU_200.DLL" ( ByVal BoosterI As Integer , ByVal Voltage As Double , ByVal Current As Double , ByVal OutFormat As Integer , ByVal VoltRef As Integer )   
Declare Sub BstISyncSet Lib "r3cu_SAU_200.DLL" ( ByVal BoosterI As Integer , ByVal SyncRef As Integer )   
'if !defined(__SAU_CONSTS_H)
Public Const uuSAU_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_SAU_UNIT As String =            "RP3_SauUnit"
'endif
'if !defined(__PSU_H)
Public Const uuPSU_H As Integer = 0
'include "r3cuPsuConsts.h"
'include "r3cuPsuTypes.h"
'ifdef __cplusplus
Declare Function PsuLibraryAttach Lib "r3cu_PSU_200.DLL" ( ) As Integer 
Declare Function PsuLibraryDetach Lib "r3cu_PSU_200.DLL" ( ) As Integer 
Declare Sub PsuDefaultsSet Lib "r3cu_PSU_200.DLL" ( ByVal vbEvent As Integer )   
Declare Function PpsuSourceConnectInterf Lib "r3cu_PSU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function PpsuSourceDisconnectInterf Lib "r3cu_PSU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function PpsuOff Lib "r3cu_PSU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function PpsuOn Lib "r3cu_PSU_200.DLL" ( ByVal InstrId As Integer , ByVal Vvalue As Double , ByVal Ivalue As Double ) As Integer 
Declare Function PpsuSourceSet Lib "r3cu_PSU_200.DLL" ( ByVal InstrId As Integer , ByVal Vvalue As Double , ByVal Ivalue As Double ) As Integer 
Declare Function PpsuSourceEnable Lib "r3cu_PSU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function PpsuSourceDisable Lib "r3cu_PSU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function PpsuReadbackV Lib "r3cu_PSU_200.DLL" ( ByVal InstrId As Integer , ByRef Vvalue As Double ) As Integer 
Declare Function PpsuReadbackI Lib "r3cu_PSU_200.DLL" ( ByVal InstrId As Integer , ByRef Ivalue As Double ) As Integer 
Declare Function PpsuGetError Lib "r3cu_PSU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function PpsuConnectUabusBmu Lib "r3cu_PSU_200.DLL" ( ByVal InstrId As Integer , ByVal RowHotP As Integer , ByVal RowHotPK As Integer , ByVal RowColdP As Integer , ByVal RowColdPK As Integer , ByVal RowHotS As Integer , ByVal RowColdS As Integer ) As Integer 
Declare Function PpsuDisconnectUabusBmu Lib "r3cu_PSU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Sub PpsOn Lib "r3cu_PSU_200.DLL" ( ByVal id As Integer , ByVal voltOut As Double , ByVal currLim As Double , ByVal timeOnOff As Integer )   
Declare Sub PpsOff Lib "r3cu_PSU_200.DLL" ( ByVal id As Integer , ByVal timeOnOff As Integer )   
Declare Sub PpsOffAll Lib "r3cu_PSU_200.DLL" ( ByVal timeOnOff As Integer )   
Declare Sub PpsReadBack Lib "r3cu_PSU_200.DLL" ( ByVal id As Integer , ByVal readbackType As Integer , ByRef Value As Double )   
Declare Sub PpsVoltageRead Lib "r3cu_PSU_200.DLL" ( ByVal id As Integer , ByRef voltOut As Double )   
Declare Sub PpsCurrentRead Lib "r3cu_PSU_200.DLL" ( ByVal id As Integer , ByRef curr As Double )   
'if !defined(__PSU_CONSTS_H)
Public Const uuPSU_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_PSU_UNIT As String =            "RP3_PsuUnit"
'endif
'if !defined(__RAC_H)
Public Const uuRAC_H As Integer = 0
'include "r3cuRacConsts.h"
'include "r3cuRacTypes.h"
'ifdef __cplusplus
Declare Function RacLibraryAttach Lib "r3cu_RAC_200.DLL" ( ) As Integer 
Declare Function RacLibraryDetach Lib "r3cu_RAC_200.DLL" ( ) As Integer 
Declare Sub RacDefaultsSet Lib "r3cu_RAC_200.DLL" ( ByVal vbEvent As Integer )   
Declare Function FpsConnectInterf Lib "r3cu_RAC_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FpsDisconnectInterf Lib "r3cu_RAC_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FpsDisconnectAllInterf Lib "r3cu_RAC_200.DLL" ( ) As Integer 
Declare Function FpsEnable Lib "r3cu_RAC_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FpsDisable Lib "r3cu_RAC_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FpsDisableAll Lib "r3cu_RAC_200.DLL" ( ) As Integer 
Declare Function FpsOn Lib "r3cu_RAC_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FpsOff Lib "r3cu_RAC_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FpsOffAll Lib "r3cu_RAC_200.DLL" ( ) As Integer 
Declare Function FpsConnectUabusBmu Lib "r3cu_RAC_200.DLL" ( ByVal InstrId As Integer , ByVal RowHotP As Integer , ByVal RowHotPK As Integer , ByVal RowColdP As Integer , ByVal RowColdPK As Integer , ByVal RowHotS As Integer , ByVal RowColdS As Integer ) As Integer 
Declare Function FpsDisconnectUabusBmu Lib "r3cu_RAC_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Sub ClearPsA Lib "r3cu_RAC_200.DLL" ( )   
Declare Sub PsAOffAll Lib "r3cu_RAC_200.DLL" ( )   
Declare Sub PsAOn Lib "r3cu_RAC_200.DLL" ( ByVal PsA As Integer )   
Declare Sub PsAOff Lib "r3cu_RAC_200.DLL" ( ByVal PsA As Integer )   
'if !defined(__RAC_CONSTS_H)
Public Const uuRAC_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_RAC_UNIT As String =            "RP3_RacUnit"
'endif
'if !defined(__SU_H)
Public Const uuSU_H As Integer = 0
'include "r3cuSuConsts.h"
'include "r3cuSuTypes.h"
'ifdef __cplusplus
Declare Function SuLibraryAttach Lib "r3cu_SU_200.DLL" ( ) As Integer 
Declare Function SuLibraryDetach Lib "r3cu_SU_200.DLL" ( ) As Integer 
Declare Function DriConnectInterf Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal vbPoint As Integer ) As Integer 
Declare Function DriDisconnectInterf Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal vbPoint As Integer ) As Integer 
Declare Function DriConnectAbus Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal HotRow As Integer , ByVal ColdRow As Integer ) As Integer 
Declare Function DriDisconnectAbus Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal HotRow As Integer , ByVal ColdRow As Integer ) As Integer 
Declare Function DriReadbackVConnectMbus Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal MbusLine As Integer , ByVal Gain As Integer ) As Integer 
Declare Function DriReadbackVDisconnectMbus Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DriReadbackIConnectMbus Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal MbusLine As Integer , ByVal Gain As Integer ) As Integer 
Declare Function DriReadbackIDisconnectMbus Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DriExtInputConnect Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal Source As Integer ) As Integer 
Declare Function DriExtInputDisconnect Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DriDisconnectAll Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DriSignalGeneratorSet Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal AmpRange As Integer , ByVal Amplitude As Double , ByVal OffRange As Integer , ByVal DCOffset As Double , ByVal IRange As Integer , ByVal WformType As Integer , ByVal Frequency As Double , ByVal OutFormat As Integer ) As Integer 
Declare Function DriSourceSet Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal Vvalue As Double , ByVal Ivalue As Double , ByVal IRange As Integer , ByVal OutMode As Integer , ByVal OutFormat As Integer , ByVal SlewRate As Integer , ByVal ByPass As Integer ) As Integer 
Declare Function DriSourceSetV2 Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal Vvalue As Double , ByVal Ivalue As Double , ByVal VRange As Integer , ByVal IRange As Integer , ByVal OutMode As Integer , ByVal OutFormat As Integer , ByVal SlewRate As Integer , ByVal ByPass As Integer ) As Integer 
Declare Function DriTimingSet Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal Flag As Integer ) As Integer 
Declare Function DriEnable Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DriDisable Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DriClear Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DriSyncSet Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal Flag As Integer , ByVal SbusLine As Integer ) As Integer 
Declare Function DriSetAdjust Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal DacId As Integer , ByVal Value As Double ) As Integer 
Declare Function DriGetAdjust Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal DacId As Integer , ByRef Value As Double ) As Integer 
Declare Function DriStoreAdjust Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DriLoadAdjust Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function GuardSet Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal SlewRate As Integer , ByVal InputRow As Integer , ByVal OutputRow As Integer , ByVal SenseRow As Integer ) As Integer 
Declare Function GuardDisable Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DriConnectUabusBmu Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal RowHotP As Integer , ByVal RowColdP As Integer , ByVal RowHotS As Integer , ByVal RowColdS As Integer ) As Integer 
Declare Function DriDisconnectUabusBmu Lib "r3cu_SU_200.DLL" ( ByVal InstrId As Integer , ByVal vbPoint As Integer ) As Integer 
Declare Sub PdrivClear Lib "r3cu_SU_200.DLL" ( ByVal Driver As Integer )   
Declare Sub PdrivEnable Lib "r3cu_SU_200.DLL" ( ByVal Driver As Integer )   
Declare Sub PdrivDisable Lib "r3cu_SU_200.DLL" ( ByVal Driver As Integer )   
Declare Sub PdrivConnectAbus Lib "r3cu_SU_200.DLL" ( ByVal Driver As Integer , ByVal vbPoint As Integer , ByVal Row As Integer )   
Declare Sub PdrivDisconnectAbus Lib "r3cu_SU_200.DLL" ( ByVal Driver As Integer , ByVal vbPoint As Integer , ByVal Row As Integer )   
Declare Sub PdrivConnectInterf Lib "r3cu_SU_200.DLL" ( ByVal Driver As Integer , ByVal vbPoint As Integer )   
Declare Sub PdrivDisconnectInterf Lib "r3cu_SU_200.DLL" ( ByVal Driver As Integer , ByVal vbPoint As Integer )   
Declare Sub PdrivDisconnectAll Lib "r3cu_SU_200.DLL" ( ByVal Driver As Integer )   
Declare Sub PdrivSourceSet Lib "r3cu_SU_200.DLL" ( ByVal Driver As Integer , ByVal OutMode As Integer , ByVal Voltage As Double , ByVal Current As Double , ByVal OutFormat As Integer , ByVal VoltRef As Integer , ByVal OutState As Integer )   
Declare Sub PdrivSyncSet Lib "r3cu_SU_200.DLL" ( ByVal Driver As Integer , ByVal SyncRef As Integer )   
Declare Sub PdrivTypeSet Lib "r3cu_SU_200.DLL" ( ByVal Driver As Integer , ByVal vbType As Integer )   
'if !defined(__SU_CONSTS_H)
Public Const uuSU_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_SU_UNIT As String =            "RP3_SuUnit"
'endif
'if !defined(__r3cuFT1000_H)
Public Const uur3cuFT1000_H As Integer = 0
'include "r3cuFt1000Consts.h"
'include "r3cuFt1000Types.h"
'ifdef __cplusplus
Declare Function FT1000LibraryAttach Lib "r3cu_FTM_200.DLL" ( ) As Integer 
Declare Function FT1000LibraryDetach Lib "r3cu_FTM_200.DLL" ( ) As Integer 
Declare Sub FT1000DefaultsSet Lib "r3cu_FTM_200.DLL" ( ByVal vbEvent As Integer )   
Declare Sub FtDriSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal Value As Double )   
Declare Sub FtDriConnect Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer )   
Declare Sub FtDriDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer )   
Declare Function DriSet Lib "r3cu_FTM_200.DLL" ( ByVal Value As Double , ByVal ChanStr As String ) As Integer 
Declare Function DriSet1 Lib "r3cu_FTM_200.DLL" ( ByVal Value As Double , ByVal Chan As Integer ) As Integer 
Declare Function MeasSeRead Lib "r3cu_FTM_200.DLL" ( ByVal Chan As Integer , ByVal Filter As Integer , ByVal Gain As Integer , ByRef Value As Double ) As Integer 
Declare Function MeasDiffRead Lib "r3cu_FTM_200.DLL" ( ByVal ChanPos As Integer , ByVal ChanNeg As Integer , ByVal Filter As Integer , ByVal Gain As Integer , ByRef Value As Double ) As Integer 
Declare Function DriConnect Lib "r3cu_FTM_200.DLL" ( ByVal ChanStr As String ) As Integer 
Declare Function DriDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal ChanStr As String ) As Integer 
'  int WINAPI pDriConnUnconn (int FlagConnect, char *ChanStr);
'  int WINAPI pDriConnUnconn1 (int FlagConnect, int Chan);
Declare Function FtDvmRmsEnable Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FtDvmRmsDisable Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FtDvmDisconnectAll Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer ) As Integer 
'  int WINAPI pFtDvmDisconnect_low (int InstrId);
Declare Function MeasSeAcquire Lib "r3cu_FTM_200.DLL" ( ByVal Chan As Integer , ByVal Filter As Integer , ByVal Gain As Integer , ByRef Value As Double , ByVal nVal As Integer ) As Integer 
Declare Function MeasDiffAcquire Lib "r3cu_FTM_200.DLL" ( ByVal ChanPos As Integer , ByVal ChanNeg As Integer , ByVal Filter As Integer , ByVal Gain As Integer , ByRef Value As Double , ByVal nVal As Integer ) As Integer 
Declare Function SyncSet Lib "r3cu_FTM_200.DLL" ( ByVal Sync As Integer , ByVal Level As Integer ) As Integer 
Declare Function FtCntClear Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer ) As Integer 
'  int pFtCntTrx (int InstrId);
Declare Function FtCntConnDisc Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal InStage As Integer , ByVal Conn As Integer , ByVal FlagConnect As Integer ) As Integer 
Declare Function FtCntConnect Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal InStage As Integer , ByVal Conn As Integer ) As Integer 
Declare Function FtCntDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal InStage As Integer , ByVal Conn As Integer ) As Integer 
Declare Function FtCntDisconnectAll Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FtCntTrigScaleSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal vbScale As Integer ) As Integer 
Declare Function FtCntTimeoutSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal Timeout As Integer ) As Integer 
'  int pFtCntProgTrx (int InstrId);
'  int pFtCnt20ProgTrx (int InstrId);
Declare Function FtCntCounterSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal Slope As Integer , ByVal Time As Integer ) As Integer 
Declare Function FtCntFreqSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal Resol As Integer , ByVal Slope As Integer ) As Integer 
Declare Function FtCntIntervalSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal Resol As Integer , ByVal Slope1 As Integer , ByVal Slope2 As Integer ) As Integer 
Declare Function FtCntAcqIntervalSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal AcqNum As Integer , ByVal Slope1 As Integer , ByVal Slope2 As Integer ) As Integer 
Declare Function FtCntPeriodSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal Resol As Integer , ByVal Slope As Integer ) As Integer 
Declare Function FtCntAcqPeriodSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal AcqNum As Integer , ByVal Slope As Integer ) As Integer 
Declare Function FtCntRatioSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal Resol As Integer , ByVal Slope1 As Integer , ByVal Slope2 As Integer ) As Integer 
Declare Function FtCntInputSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal InStage As Integer , ByVal TrigThr As Double , ByVal Filter As Integer , ByVal Attenuator As Integer , ByVal Coupling As Integer ) As Integer 
Declare Function FtCntAcq Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByRef MeasValue As Double , ByRef MeasType As Integer , ByVal AcqId As Integer ) As Integer 
'  int WINAPI pFtCntCalib (int InstrId, int Param, int Dac);
'  int WINAPI pFtCntE2promRW (int InstrId, int writeFlag, int *VrefL, int *VrefH, long *Freq);
Declare Function FtCntSyncSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iSelCh As Integer , ByVal iSlopeSync As Integer ) As Integer 
Declare Function FtCntSyncRead Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByRef MeasValue As Double , ByRef MeasType As Integer , ByVal AcqId As Integer ) As Integer 
Declare Function FtCntClockSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal vbScale As Integer ) As Integer 
Declare Function FtCntDutyCycleSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal Resol As Integer , ByVal Slope As Integer ) As Integer 
Declare Function FtCntAcqDutyCycleSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal AcqNum As Integer , ByVal Slope As Integer ) As Integer 
Declare Function FtCntDutyCycleAcq Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByRef DutyCycleValue As Double , ByRef TonValue As Double , ByRef PeriodValue As Double , ByRef MeasType As Integer , ByVal AcqId As Integer ) As Integer 
Declare Function FtCntDutyCycleSyncRead Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByRef DutyCycleValue As Double , ByRef TonValue As Double , ByRef PeriodValue As Double , ByRef MeasType As Integer , ByVal AcqId As Integer ) As Integer 
Declare Function FtAcIMeasGainSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iChan As Integer , ByVal iGain As Integer ) As Integer 
Declare Function FtDcGenConnect Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FtDcGenDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FtDcGenEnable Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FtDcGenDisable Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function FtDcGenVoltageSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal dVoltage As Double ) As Integer 
Declare Function FtDcGenCurrentSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal dCurrent As Double ) As Integer 
Declare Function DoutSet Lib "r3cu_FTM_200.DLL" ( ByVal Level As Integer , ByVal ChanStr As String ) As Integer 
Declare Function DoutSet1 Lib "r3cu_FTM_200.DLL" ( ByVal Level As Integer , ByVal Chan As Integer ) As Integer 
Declare Function DinpReadAll Lib "r3cu_FTM_200.DLL" ( ) As Integer 
Declare Function DinpRead Lib "r3cu_FTM_200.DLL" ( ByRef Level As Integer , ByVal Chan As Integer ) As Integer 
Declare Function DinpSyncSet Lib "r3cu_FTM_200.DLL" ( ByVal Trigger As Integer , ByVal Chan As Integer , ByVal Level As Integer ) As Integer 
Declare Function DinpMeasAcq Lib "r3cu_FTM_200.DLL" ( ByRef ChannelList As Integer , ByVal Strobe As Integer , ByVal NumAcq As Integer ) As Integer 
Declare Function DinpTimeReadStart Lib "r3cu_FTM_200.DLL" ( ByVal Strobe As Integer ) As Integer 
Declare Function DinpTimeReadStop Lib "r3cu_FTM_200.DLL" ( ByRef Cnt1 As Double , ByRef Cnt2 As Double ) As Integer 
Declare Function DinpMeasRead Lib "r3cu_FTM_200.DLL" ( ByVal Chan As Integer , ByRef AcquiredList As Integer ) As Integer 
Declare Function DinpTest Lib "r3cu_FTM_200.DLL" ( ByRef Level As Integer , ByVal Chan As Integer ) As Integer 
Declare Function DinpClear Lib "r3cu_FTM_200.DLL" ( ) As Integer 
Declare Function DoutClear Lib "r3cu_FTM_200.DLL" ( ) As Integer 
Declare Function FtDvmConnect Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iHotPoint As Integer , ByVal iColdPoint As Integer ) As Integer 
Declare Function FtDvmDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtDvmSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iRange As Integer , ByVal iFilter As Integer ) As Integer 
Declare Function FtDvmSyncSet2 Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iSync As Integer , ByVal iLevel As Integer ) As Integer 
Declare Function FtDvmStrobeSet2 Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iSync As Integer , ByVal iLevel As Integer ) As Integer 
Declare Function FtDvmSyncSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iSync As Integer , ByVal iLevel As Integer , ByVal iStrobe As Integer , ByVal iValue As Integer ) As Integer 
Declare Function FtDvmMeas Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByRef lpdValue As Double ) As Integer 
Declare Function FtDvmComp Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByRef lpiResult As Integer ) As Integer 
Declare Function FtDvmMeasComp Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByRef lpdValue As Double , ByRef lpiResult As Integer ) As Integer 
Declare Function FtDvmMeasAcq Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iValuesQty As Integer ) As Integer 
Declare Function FtDvmMeasAcqStart Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iValuesQty As Integer ) As Integer 
Declare Function FtDvmMeasAcqWaitEnd Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtDvmCompAcq Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iValuesQty As Integer ) As Integer 
Declare Function FtDvmMeasCompAcq Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iValuesQty As Integer ) As Integer 
Declare Function FtDvmMeasRead Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByRef lpdValueList As Double , ByRef lpiResultList As Integer , ByRef lpiValuesQty As Integer ) As Integer 
Declare Function FtDvmCompRead Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByRef lpiResultList As Integer , ByRef lpiValuesQty As Integer ) As Integer 
Declare Function FtDvmCompSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal dThrH As Double , ByVal dThrL As Double ) As Integer 
Declare Function FtDvmTimeoutSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iTimeout As Integer ) As Integer 
Declare Function FtDvmMeasAcqPar Lib "r3cu_FTM_200.DLL" ( ByRef iInstrIdList As Integer , ByVal iValuesQty As Integer ) As Integer 
Declare Function FtDvmCompAcqPar Lib "r3cu_FTM_200.DLL" ( ByRef iInstrIdList As Integer , ByVal iValuesQty As Integer ) As Integer 
Declare Function FtDvmMeasCompAcqPar Lib "r3cu_FTM_200.DLL" ( ByRef iInstrIdList As Integer , ByVal iValuesQty As Integer ) As Integer 
Declare Function FtDvmCalibSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iEnableFlag As Integer ) As Integer 
Declare Function FtDvmCalibDacSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iOut As Integer , ByVal iValue As Integer ) As Integer 
Declare Function FtDvmEepromWriteWord Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iAddress As Integer , ByVal iValue As Integer ) As Integer 
Declare Function FtDvmEepromWriteDouble Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iAddress As Integer , ByVal dValue As Double ) As Integer 
Declare Function FtDvmEepromReadWord Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iAddress As Integer , ByRef iValue As Integer ) As Integer 
Declare Function FtDvmEepromReadDouble Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iAddress As Integer , ByRef dValue As Double ) As Integer 
Declare Function EepromWriteVerify Lib "r3cu_FTM_200.DLL" ( ByVal FileName As String , ByVal StartAddress As Integer ) As Integer 
Declare Function EepromRead Lib "r3cu_FTM_200.DLL" ( ByVal FileName As String , ByVal nBit As Integer ) As Integer 
Declare Function FtCurrMeasGainSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal byGain As Short ) As Integer 
Declare Function FtCurrMeasConnect Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtCurrMeasDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtCurrAcCoupEnable Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtCurrAcCoupDisable Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtCurrRmsEnable Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtCurrRmsDisable Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function HvChannelClear Lib "r3cu_FTM_200.DLL" ( ) As Integer 
Declare Function HvChannelConnect Lib "r3cu_FTM_200.DLL" ( ByVal ChanStr As String , ByVal Row As Integer ) As Integer 
Declare Function HvChannelDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal ChanStr As String , ByVal Row As Integer ) As Integer 
'  int WINAPI pHvChanConnUnconn (int Connect, char *ChanStr, int Row);
Declare Function HvChannelConnect1 Lib "r3cu_FTM_200.DLL" ( ByVal Chan As Integer , ByVal Row As Integer ) As Integer 
Declare Function HvChannelDisconnect1 Lib "r3cu_FTM_200.DLL" ( ByVal Chan As Integer , ByVal Row As Integer ) As Integer 
'  int WINAPI pHvChanConnUnconn1 (int Connect, int Chan, int Row);
Declare Function HvChannelDisconnectAll Lib "r3cu_FTM_200.DLL" ( ) As Integer 
'  int pFtHvDcGenTrx (int HvGen, BOOL FlagSet);
Declare Function FtHvDcGenClear Lib "r3cu_FTM_200.DLL" ( ByVal HvGen As Integer ) As Integer 
Declare Function FtHvDcGenEnable Lib "r3cu_FTM_200.DLL" ( ByVal HvGen As Integer ) As Integer 
Declare Function FtHvDcGenDisable Lib "r3cu_FTM_200.DLL" ( ByVal HvGen As Integer ) As Integer 
Declare Function FtHvDcGenConnect Lib "r3cu_FTM_200.DLL" ( ByVal HvGen As Integer ) As Integer 
Declare Function FtHvDcGenDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal HvGen As Integer ) As Integer 
Declare Function FtHvDcGenSet Lib "r3cu_FTM_200.DLL" ( ByVal HvGen As Integer , ByVal Voltage As Double ) As Integer 
Declare Sub HvGenEnable Lib "r3cu_FTM_200.DLL" ( ByVal HvGen As Integer )   
Declare Sub HvGenDisable Lib "r3cu_FTM_200.DLL" ( ByVal HvGen As Integer )   
Declare Sub HvGenConnect Lib "r3cu_FTM_200.DLL" ( ByVal HvGen As Integer )   
Declare Sub HvGenDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal HvGen As Integer )   
Declare Sub HvGenSet Lib "r3cu_FTM_200.DLL" ( ByVal HvGen As Integer , ByVal Voltage As Double )   
Declare Function FtMuxConnect Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iChan As Integer ) As Integer 
Declare Function FtMuxDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtFpsOn Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer ) As Integer 
Declare Function PsOn Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer ) As Integer 
Declare Function FtFpsOff Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer ) As Integer 
Declare Function PsOff Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer ) As Integer 
Declare Function FtPpsClear Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer , ByVal ResetAll As Integer ) As Integer 
Declare Function FtPpsConnect Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer ) As Integer 
Declare Function FtPpsDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer ) As Integer 
Declare Function FtPpsConnectMon Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer ) As Integer 
Declare Function FtPpsDisconnectMon Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer ) As Integer 
Declare Function FtPpsGetOpMode Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer , ByRef OpMode As Integer ) As Integer 
Declare Function FtPpsOn Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer ) As Integer 
Declare Function FtPpsOff Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer ) As Integer 
Declare Function FtPpsVoltageSet Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer , ByVal Value As Double ) As Integer 
Declare Function FtPpsCurrentSet Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer , ByVal Value As Double ) As Integer 
'  int WINAPI pFtPpsE2promRW (int Ps, int writeFlag, 'struct Pps_E2prom *E2prom);
Declare Function FtPpsWriteEEprom Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer , ByVal Address As Integer , ByVal Value As Integer ) As Integer 
Declare Function FtPpsReadEEprom Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer , ByVal Address As Integer , ByRef Value As Integer ) As Integer 
Declare Function FtPpsWriteEEpromDouble Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer , ByVal Address As Integer , ByVal Value As Double ) As Integer 
Declare Function FtPpsReadEEpromDouble Lib "r3cu_FTM_200.DLL" ( ByVal Ps As Integer , ByVal Address As Integer , ByRef Value As Double ) As Integer 
Declare Function RelaySet Lib "r3cu_FTM_200.DLL" ( ByVal FlagClose As Integer , ByVal RelayStr As String ) As Integer 
Declare Function RelaySet1 Lib "r3cu_FTM_200.DLL" ( ByVal FlagClose As Integer , ByVal Relay As Integer ) As Integer 
Declare Function RelayClear Lib "r3cu_FTM_200.DLL" ( ) As Integer 
Declare Function FtPresSet Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal GroupId As Integer , ByVal Value As Double ) As Integer 
Declare Function FtPresConnect Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal GroupId As Integer ) As Integer 
Declare Function FtPresDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal GroupId As Integer ) As Integer 
Declare Function FtPresWrite Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal Dati As String , ByVal vbLen As Integer ) As Integer 
Declare Function FtPresRead Lib "r3cu_FTM_200.DLL" ( ByVal InstrId As Integer , ByVal Dati As String , ByVal vbLen As Integer ) As Integer 
Declare Function FtReslConnect Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtReslDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtReslSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal dResProg As Double ) As Integer 
Declare Function FtReslWriteEEProm Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iAddress As Integer , ByVal dValue As Double ) As Integer 
Declare Function FtReslReadEEProm Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iAddress As Integer , ByRef dValue As Double ) As Integer 
Declare Function FtSCondRmsEnable Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtSCondRmsDisable Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtSCondGainSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iGain As Integer ) As Integer 
Declare Function FtSCondFilterSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iFilter As Integer ) As Integer 
Declare Function FtSCondConnect Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtSCondDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function Wait Lib "r3cu_FTM_200.DLL" ( ByVal mSec As Integer ) As Integer 
Declare Function TimeInterval Lib "r3cu_FTM_200.DLL" ( ByRef uSec As Integer , ByVal StartChan As Integer , ByVal StartLevel As Integer , ByVal EndChan As Integer , ByVal EndLevel As Integer , ByVal TimeOut As Integer ) As Integer 
Declare Function UserWrite Lib "r3cu_FTM_200.DLL" ( ByVal Add As Integer , ByVal vbData As Integer ) As Integer 
Declare Function UserRead Lib "r3cu_FTM_200.DLL" ( ByVal Add As Integer , ByRef vbData As Integer ) As Integer 
Declare Function FtWfgConnect Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iConnect As Integer ) As Integer 
Declare Function FtWfgDisconnect Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer ) As Integer 
Declare Function FtWfgSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal dFrequency As Double , ByVal dAmplitude As Double , ByVal dDutyCycle As Double , ByVal dOffset As Double , ByVal iType As Integer ) As Integer 
Declare Function FtWfgWriteEEProm Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iAddress As Integer , ByVal dValue As Single ) As Integer 
Declare Function FtWfgReadEEProm Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iAddress As Integer , ByRef dValue As Single ) As Integer 
Declare Function FtWfgCalSet Lib "r3cu_FTM_200.DLL" ( ByVal iInstrId As Integer , ByVal iOutDac As Integer , ByVal dVdac As Double , ByVal iRange As Integer , ByVal iType As Integer ) As Integer 
'if !defined(__FT1000_CONSTS_H)
Public Const uuFT1000_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_FT1000_UNIT As String =            "RP3_FT1000Unit"
'endif
'if !defined(__YAPROCO_H)
Public Const uuYAPROCO_H As Integer = 0
'include "r3cuProcoConsts.h"
'include "r3cuProcoTypes.h"
'ifdef __cplusplus
Declare Function ProcoLibraryAttach Lib "r3cu_proco_200.DLL" ( ) As Integer 
Declare Function ProcoLibraryDetach Lib "r3cu_proco_200.DLL" ( ) As Integer 
Declare Function PrbConnectUabusAbus Lib "r3cu_proco_200.DLL" ( ByVal ProbeGroup As Integer ) As Integer 
Declare Function PrbDisconnectUabusAbus Lib "r3cu_proco_200.DLL" ( ByVal ProbeGroup As Integer ) As Integer 
Declare Function PrbUabusModeSet Lib "r3cu_proco_200.DLL" ( ByVal ProbeGroup As Integer , ByVal ProbeRow As Integer ) As Integer 
Declare Function PrbChConnectUabus Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeChan As Integer , ByVal ProbeRow As Integer ) As Integer 
Declare Function PrbChDisconnectUabus Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeChan As Integer , ByVal ProbeRow As Integer ) As Integer 
Declare Function PrbChConnect Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeChan As Integer , ByVal ProbeRow As Integer ) As Integer 
Declare Function PrbChDisconnect Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeChan As Integer , ByVal ProbeRow As Integer ) As Integer 
Declare Function PrbAChOutConnectUabus Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeRow As Integer ) As Integer 
Declare Function PrbAChOutDisconnectUabus Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeRow As Integer ) As Integer 
Declare Function PrbAChInConnectUabus Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeRow As Integer ) As Integer 
Declare Function PrbAChInDisconnectUabus Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeRow As Integer ) As Integer 
Declare Function PrbCurrScaleSet Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeScale As Integer ) As Integer 
Declare Function PrbAChInMeasType Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeMeasType As Integer ) As Integer 
Declare Function PrbAChInGain Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeGain As Integer ) As Integer 
Declare Function PrbAChOutGain Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeGain As Integer ) As Integer 
Declare Function PrbDChConnect Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbDChDisconnect Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbDChFloTest Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByRef ProbeResult As Integer ) As Integer 
Declare Function PrbDChLevelSet Lib "r3cu_proco_200.DLL" ( ByVal ProbeVh As Double ) As Integer 
Declare Function PrbDChImpSet Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeOutImp As Double ) As Integer 
Declare Function PrbDChSensorSet Lib "r3cu_proco_200.DLL" ( ByVal ProbeVth As Double ) As Integer 
Declare Function PrbDChVppLevelSet Lib "r3cu_proco_200.DLL" ( ByVal ProbeVhh As Double ) As Integer 
Declare Function PrbChConnectInterf Lib "r3cu_proco_200.DLL" ( ByVal ProbeGroup As Integer ) As Integer 
Declare Function PrbChConnectInternal Lib "r3cu_proco_200.DLL" ( ByVal ProbeGroup As Integer ) As Integer 
Declare Function PrbFreqDivDiagEnable Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbFreqDivDiagDisable Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbFreqDivSet Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbNumDiv As Integer ) As Integer 
Declare Function PrbFreqDivClear Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbRlcmConnect Lib "r3cu_proco_200.DLL" ( ByVal ProbeHot As Integer , ByVal ProbeCold As Integer ) As Integer 
Declare Function PrbRlcmDisconnect Lib "r3cu_proco_200.DLL" ( ) As Integer 
Declare Function PrbRlcmScaleSet Lib "r3cu_proco_200.DLL" ( ByVal ProbeScale As Integer ) As Integer 
Declare Function PrbRlcmMeas Lib "r3cu_proco_200.DLL" ( ByVal ProbeMeasType As Integer , ByVal ProbeCircuitType As Integer , ByRef ProbeMeasValue As Double ) As Integer 
Declare Function PrbGndConnect Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbGndDisconnect Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbShConnect Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbShDisconnect Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbEscanEnable Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbEscanDisable Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbTempRead Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByRef TempValue As Double ) As Integer 
Declare Function PrbDiagModeEnable Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbDiagModeDisable Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbActiveShieldEnable Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbActiveShieldDisable Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbShoShieldEnable Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbShoShieldDisable Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer ) As Integer 
Declare Function PrbLcsSensorMeas Lib "r3cu_proco_200.DLL" ( ByVal pProbeId As Integer , ByVal pTime As Double , ByVal pColorGain As Integer , ByVal pLuxGain As Integer ) As Integer 
Declare Function PrbLcsColorReadHSL Lib "r3cu_proco_200.DLL" ( ByVal pProbeId As Integer , ByRef pHue As Double , ByRef pSat As Double , ByRef pLum As Double ) As Integer 
Declare Function PrbLcsColorReadRGB Lib "r3cu_proco_200.DLL" ( ByVal pProbeId As Integer , ByRef pRed As Integer , ByRef pGreen As Integer , ByRef pBlue As Integer ) As Integer 
Declare Function PrbLcsLuxRead Lib "r3cu_proco_200.DLL" ( ByVal pProbeId As Integer , ByRef pLux As Double , ByRef pLuxIR As Double ) As Integer 
Declare Function PrbDChSet Lib "r3cu_proco_200.DLL" ( ByVal pProbeId As Integer , ByVal pObpCh As Short , ByVal pProbeOutImp As Double ) As Integer 
Declare Function PrbDChStuckSet Lib "r3cu_proco_200.DLL" ( ByVal pProbeId As Integer , ByVal pLogicLevel As Integer ) As Integer 
Declare Function PrbDChLevelRead Lib "r3cu_proco_200.DLL" ( ByVal pProbeId As Integer , ByRef pLogicLevel As Integer ) As Integer 
Declare Function PrbDChReset Lib "r3cu_proco_200.DLL" ( ByVal pProbeId As Integer ) As Integer 
Declare Function PrbFreqDivSetRow Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeNumDiv As Integer , ByVal ProbeRow As Integer ) As Integer 
Declare Function PrbEscanConnectUabus Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeRow As Integer ) As Integer 
Declare Function PrbEscanDisconnectUabus Lib "r3cu_proco_200.DLL" ( ByVal ProbeId As Integer , ByVal ProbeRow As Integer ) As Integer 
'if !defined(__YAPROCO_CONSTS_H)
Public Const uuYAPROCO_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_YAPROCO_UNIT As String =            "RP3_YaProcoUnit"
'endif
'if !defined(__PW_H)
Public Const uuPW_H As Integer = 0
'include "r3cuPwConsts.h"
'include "r3cuPwTypes.h"
'ifdef __cplusplus
Declare Function PwLibraryAttach Lib "r3cu_POWER_200.DLL" ( ) As Integer 
Declare Function PwLibraryDetach Lib "r3cu_POWER_200.DLL" ( ) As Integer 
Declare Sub PwDefaultsSet Lib "r3cu_POWER_200.DLL" ( ByVal vbEvent As Integer )   
Declare Function AcGenConnect Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function AcGenOutConnect Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal AcGenOut As Integer ) As Integer 
Declare Function AcGenOutDisconnect Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal AcGenOut As Integer ) As Integer 
Declare Function AcGenDisconnect Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DcGenConnect Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DcGenOutConnect Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal AcGenOut As Integer ) As Integer 
Declare Function DcGenOutDisconnect Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal AcGenOut As Integer ) As Integer 
Declare Function DcGenDisconnect Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function PmxConnect Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal Section As Integer , ByRef Channel As Short ) As Integer 
Declare Function PmxDisconnect Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal Section As Integer , ByRef Channel As Short ) As Integer 
Declare Function PmxDisconnectAll Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function AcGenSourceSet Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal VValue As Double , ByVal VRange As Integer , ByVal IValue As Double , ByVal ILimit As Double , ByVal Frequency As Double ) As Integer 
Declare Function AcGenSourceSetEx Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal VValue As Double , ByVal VRange As Integer , ByVal IValue As Double , ByVal ILimit As Double , ByVal Frequency As Double , ByVal Vdc As Double , ByVal Phase As Double ) As Integer 
Declare Function AcGenEnable Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function AcGenDisable Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function AcGenOn Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal VValue As Double , ByVal VRange As Integer , ByVal IValue As Double , ByVal ILimit As Double , ByVal Frequency As Double ) As Integer 
Declare Function AcGenOnEx Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal VValue As Double , ByVal VRange As Integer , ByVal IValue As Double , ByVal ILimit As Double , ByVal Frequency As Double , ByVal Vdc As Double , ByVal Phase As Double ) As Integer 
Declare Function AcGenOutOn Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal AcGenOut As Integer , ByVal VValue As Double , ByVal VRange As Integer , ByVal IValue As Double , ByVal ILimit As Double , ByVal Frequency As Double ) As Integer 
Declare Function AcGenOutOnEx Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal Out As Integer , ByVal VValue As Double , ByVal VRange As Integer , ByVal IValue As Double , ByVal ILimit As Double , ByVal Frequency As Double , ByVal Vdc As Double , ByVal Phase As Double ) As Integer 
Declare Function AcGenOff Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function AcGenOutOff Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal AcGenOut As Integer ) As Integer 
Declare Function ReadAcGenDvm Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal vbType As Integer , ByRef Value As Double ) As Integer 
Declare Function DcGenSourceSet Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal VValue As Double , ByVal IValue As Double ) As Integer 
Declare Function DcGenEnable Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DcGenDisable Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DcGenOn Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal VValue As Double , ByVal IValue As Double ) As Integer 
Declare Function DcGenOutOn Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal AcGenOut As Integer , ByVal VValue As Double , ByVal IValue As Double ) As Integer 
Declare Function DcGenOff Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DcGenOutOff Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal AcGenOut As Integer ) As Integer 
Declare Function ReadDcGenDvm Lib "r3cu_POWER_200.DLL" ( ByVal InstrId As Integer , ByVal vbType As Integer , ByRef Value As Double ) As Integer 
'if !defined(__PW_CONSTS_H)
Public Const uuPW_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_PW_UNIT As String =            "RP3_PwUnit"
'endif
'if !defined(__BSCAN_H)
Public Const uuBSCAN_H As Integer = 0
'include "r3cuBscanConsts.h"
'include "r3cuBscanTypes.h"
'ifdef __cplusplus
Declare Function BscanLibraryAttach Lib "r3cu_BSCAN_200.DLL" ( ) As Integer 
Declare Function BscanLibraryDetach Lib "r3cu_BSCAN_200.DLL" ( ) As Integer 
Declare Sub BscanDefaultsSet Lib "r3cu_BSCAN_200.DLL" ( ByVal vbEvent As Integer )   
Declare Function BscanTapConnect Lib "r3cu_BSCAN_200.DLL" ( ByVal ChainName As String , ByVal TpTCK As Integer , ByVal TpTMS As Integer , ByVal TpTDI As Integer , ByVal TpTDO As Integer , ByVal TpTRST As Integer ) As Integer 
Declare Function BscanTapOpen Lib "r3cu_BSCAN_200.DLL" ( ByVal ChainName As String , ByVal MaxFrequency As Integer ) As Integer 
Declare Function BscanTapTest Lib "r3cu_BSCAN_200.DLL" ( ByVal ChainName As String ) As Integer 
Declare Function BscanIdCodeTest Lib "r3cu_BSCAN_200.DLL" ( ByVal ChainName As String ) As Integer 
Declare Function BscanTapPinsLevelSet Lib "r3cu_BSCAN_200.DLL" ( ByVal ChainName As String , ByVal DrwRef As String , ByVal PinNameArray As String , ByRef PinLevelArray As Integer ) As Integer 
Declare Function BscanExTest Lib "r3cu_BSCAN_200.DLL" ( ByVal ChainName As String ) As Integer 
Declare Function BscanTapPinsLevelRead Lib "r3cu_BSCAN_200.DLL" ( ByVal ChainName As String , ByVal DrwRef As String , ByVal PinNameArray As String , ByRef PinLevelArray As Integer ) As Integer 
Declare Function BscanTapClose Lib "r3cu_BSCAN_200.DLL" ( ByVal ChainName As String ) As Integer 
'if !defined(__BSCAN_CONSTS_H)
Public Const uuBSCAN_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_BSCAN_UNIT As String =            "RP3_BscanUnit"
'endif
'if !defined(__AUTO_H)
Public Const uuAUTO_H As Integer = 0
'include "r3cuAutoConsts.h"
'include "r3cuAutoTypes.h"
'ifdef __cplusplus
Declare Function AutoLibraryAttach Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutoLibraryDetach Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Sub AutoDefaultsSet Lib "r3cu_AUTO_200.DLL" ( ByVal vbEvent As Integer )   
Declare Function ReceiverOpenComm Lib "r3cu_AUTO_200.DLL" ( ByVal RecType As Integer , ByVal RecPort As Integer , ByVal Info As String ) As Integer 
Declare Function ReceiverOpenCommV2 Lib "r3cu_AUTO_200.DLL" ( ByVal RecType As Integer , ByVal RecPort As Integer , ByVal Info As String , ByVal pHomingFlag As Integer ) As Integer 
Declare Function ReceiverCloseComm Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ReceiverIsOpenComm Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ReceiverSystemSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ReceiverTprjSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ReceiverRampTprjSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Sub AutomationModeSetVar Lib "r3cu_AUTO_200.DLL" ( ByVal Mode As Integer )   
Declare Function AutomationModeSet Lib "r3cu_AUTO_200.DLL" ( ByVal Mode As Integer ) As Integer 
Declare Function AutomationModeGet Lib "r3cu_AUTO_200.DLL" ( ByRef Mode As Integer ) As Integer 
Declare Function AutomationInfoRead Lib "r3cu_AUTO_200.DLL" ( ByVal AutoInfo As String ) As Integer 
Declare Function ReceiverHoming Lib "r3cu_AUTO_200.DLL" ( ByVal FastFlag As Integer ) As Integer 
Declare Function ReceiverMoveUp Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ReceiverMoveDown Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ReceiverMoveMiddle Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ReceiverMoveRDown Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ReceiverMoveRMiddle Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ConveyorMoveIn Lib "r3cu_AUTO_200.DLL" ( ByVal SyncFlag As Integer ) As Integer 
Declare Function ConveyorMoveOut Lib "r3cu_AUTO_200.DLL" ( ByVal SyncFlag As Integer ) As Integer 
Declare Function ReceiverFingerAdj Lib "r3cu_AUTO_200.DLL" ( ByVal Enable As Integer , ByVal TxNow As Integer ) As Integer 
Declare Function ReceiverPositionRead Lib "r3cu_AUTO_200.DLL" ( ByRef Position As Integer ) As Integer 
Declare Function ReceiverPositionSet Lib "r3cu_AUTO_200.DLL" ( ByVal Position As Integer ) As Integer 
Declare Function ConveyorBypassSet Lib "r3cu_AUTO_200.DLL" ( ByVal BypassEnabled As Integer ) As Integer 
Declare Function ConveyorContinue Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationPolling Lib "r3cu_AUTO_200.DLL" ( ByRef Head As Integer ) As Integer 
Declare Function AutomationPollingDelete Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationHandleUnexpectedCode Lib "r3cu_AUTO_200.DLL" ( ByVal lpcMsgRead As String , ByRef lpbRetry As Integer ) As Integer 
Declare Function ReceiverSendSwActive Lib "r3cu_AUTO_200.DLL" ( ByVal Active As Integer ) As Integer 
Declare Function ReceiverSendSwStatus Lib "r3cu_AUTO_200.DLL" ( ByVal Head As Integer , ByVal Status As Integer ) As Integer 
Declare Function ReceiverSetStageStatus Lib "r3cu_AUTO_200.DLL" ( ByVal Head As Integer , ByVal Stage As Integer , ByVal Status As Integer ) As Integer 
Declare Function AutomationSendSnReadResult Lib "r3cu_AUTO_200.DLL" ( ByVal Head As Integer , ByVal SnReadResult As Integer ) As Integer 
Declare Function ReceiverSendStart Lib "r3cu_AUTO_200.DLL" ( ByRef StartAllowed As Integer ) As Integer 
Declare Function ReceiverSendStop Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ReceiverSendStopNoSync Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationSetBuzzer Lib "r3cu_AUTO_200.DLL" ( ByVal BuzMode As Integer ) As Integer 
Declare Function AutomationSetSemaphore Lib "r3cu_AUTO_200.DLL" ( ByVal Light As Integer , ByVal Enable As Integer ) As Integer 
Declare Function AutomationSetSemaphoreV2 Lib "r3cu_AUTO_200.DLL" ( ByVal pBuzzerMode As Integer , ByVal pGreenMode As Integer , ByVal pOrangeMode As Integer , ByVal pRedMode As Integer , ByVal pBlueMode As Integer ) As Integer 
Declare Function AutomationSetBcrEnable Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationSetBcrEnableV2 Lib "r3cu_AUTO_200.DLL" ( ByVal pEnable As Integer ) As Integer 
'  BOOL WINAPI AutomationRegistersWrite (UCHAR *AutoRegs, WORD RegsSize);
'  BOOL WINAPI AutomationRegistersRead (UCHAR *AutoRegs, WORD *RegsSize);
'  BOOL WINAPI AutomationRegistersReadV2 (BOOL FlagInput, UCHAR *AutoRegs, WORD *RegsSize);
Declare Function AdapterLock Lib "r3cu_AUTO_200.DLL" ( ByVal LockFlag As Integer ) As Integer 
Declare Function AutomationDoorSet Lib "r3cu_AUTO_200.DLL" ( ByVal OpenFlag As Integer ) As Integer 
'  BOOL WINAPI AutomationDoorSetV2 (UCHAR Flag);
Declare Function AutomationDebugIORead Lib "r3cu_AUTO_200.DLL" ( ByVal FlagInput As Integer , ByVal iAddress As Integer , ByRef iValue As Integer ) As Integer 
Declare Function AutomationDebugWriteReadMessageEcho Lib "r3cu_AUTO_200.DLL" ( ByVal CanId As Integer , ByVal MsgWrite As String , ByVal MsgWriteLen As Integer , ByVal MsgRead As String , ByVal MsgReadLen As Integer ) As Integer 
Declare Function AutomationDebugExternalIORead Lib "r3cu_AUTO_200.DLL" ( ByVal CanId As Integer , ByVal FlagInput As Integer , ByVal iAddress As Integer , ByRef iValue As Integer ) As Integer 
Declare Function AutomationDebugOutputWrite Lib "r3cu_AUTO_200.DLL" ( ByVal iAddress As Integer , ByVal iValue As Integer ) As Integer 
Declare Function AutomationDebugExternalOutputWrite Lib "r3cu_AUTO_200.DLL" ( ByVal CanId As Integer , ByVal iAddress As Integer , ByVal iValue As Integer ) As Integer 
Declare Function AutomationSwErrorSet Lib "r3cu_AUTO_200.DLL" ( ByVal ErrType As Integer ) As Integer 
Declare Function AutomationSwErrorWithoutBeepSet Lib "r3cu_AUTO_200.DLL" ( ByVal ErrType As Integer ) As Integer 
Declare Function AutomationMaintenanceSet Lib "r3cu_AUTO_200.DLL" ( ByVal MaintEnabled As Integer ) As Integer 
Declare Function AutomationBcrSet Lib "r3cu_AUTO_200.DLL" ( ByVal BcrEnabled As Integer ) As Integer 
Declare Function ConveyorLineHoming Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ConveyorLineHomingStatusRead Lib "r3cu_AUTO_200.DLL" ( ByRef pHomingDone As Integer ) As Integer 
Declare Function ConveyorHoming Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ConveyorPositionRead Lib "r3cu_AUTO_200.DLL" ( ByRef Position As Integer ) As Integer 
Declare Function ConveyorPositionSet Lib "r3cu_AUTO_200.DLL" ( ByVal Position As Integer , ByVal IsBoardStretch As Integer ) As Integer 
Declare Function ConveyorStartPositionSet Lib "r3cu_AUTO_200.DLL" ( ByVal Position As Integer ) As Integer 
Declare Function ConveyorMotorMove Lib "r3cu_AUTO_200.DLL" ( ByVal Start As Integer , ByVal Section As Integer , ByVal Direction As Integer ) As Integer 
Declare Function ConveyorStepMoveLow Lib "r3cu_AUTO_200.DLL" ( ByVal StepQty As Integer ) As Integer 
Declare Function AutomationSetRetrigger Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationSnMissing Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationSkipBoard Lib "r3cu_AUTO_200.DLL" ( ByVal SkipReason As Integer ) As Integer 
Declare Function AutomationLastBoard Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationSendWizardCommand Lib "r3cu_AUTO_200.DLL" ( ByVal Cmd As String , ByVal CmdLen As Integer ) As Integer 
Declare Function AutomationFixtureHoldProcedure Lib "r3cu_AUTO_200.DLL" ( ByVal DownFlag As Integer ) As Integer 
Declare Function AutomationDetailedStatusEnable Lib "r3cu_AUTO_200.DLL" ( ByVal EnableFlag As Integer ) As Integer 
Declare Function AutomationDebugCycle Lib "r3cu_AUTO_200.DLL" ( ByVal CycleId As Integer , ByVal CycleQty As Integer , ByVal Param1 As Integer , ByVal Param2 As Integer ) As Integer 
Declare Function AutomationPause Lib "r3cu_AUTO_200.DLL" ( ByVal FlagPause As Integer ) As Integer 
'  BOOL WINAPI AutomationPauseV2 (BOOL FlagPause, UCHAR Direction);
'  BOOL WINAPI AutomationSnSkipBoard (UCHAR pSkipBoardResult);
Declare Function AutomationHomingStatusSet Lib "r3cu_AUTO_200.DLL" ( ByVal pHomingRunning As Integer , ByVal pHomingOk As Integer ) As Integer 
'  BOOL WINAPI AutomationLightStatusSet (long pIllumId, typAUTOLIGHTSTATUS *pLightStatus);
'  BOOL WINAPI AutomationLightStatusGet (long pIllumId, typAUTOLIGHTSTATUS *pLightStatus);
Declare Function AutomationCapabilityGet Lib "r3cu_AUTO_200.DLL" ( ByVal pFeatureId As Integer ) As Integer 
Declare Function AutomationGetCode Lib "r3cu_AUTO_200.DLL" ( ByVal pCodeId As Integer , ByRef pCodeData As Integer ) As Integer 
Declare Function AutomationWizardSensorRead Lib "r3cu_AUTO_200.DLL" ( ByRef pSensorBmp As Integer ) As Integer 
Declare Function FomSystemSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function FomTprjSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationRslLifterMoveEnable Lib "r3cu_AUTO_200.DLL" ( ByVal EnableFlag As Integer , ByVal UpFlag As Integer ) As Integer 
Declare Function AutomationBtuSensorSetToAprobe Lib "r3cu_AUTO_200.DLL" ( ByVal pHead As Integer ) As Integer 
Declare Function ConveyorMoveInDuringTest Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ConveyorMoveOutDuringTest Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function FomOverturn Lib "r3cu_AUTO_200.DLL" ( ByVal pHead As Integer , ByVal pMode As Integer ) As Integer 
Declare Function AutomationParkProbeForDoorOpen Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationWakeupFromPowerSave Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
'  BOOL WINAPI AutomationAllowPowerSave (UCHAR pPowerSaveStatus, BOOL pAirSavePossible);
Declare Function ExtReceiverIdle Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ExtReceiverInputRead Lib "r3cu_AUTO_200.DLL" ( ByVal iInputId As Integer , ByRef iValue As Integer ) As Integer 
Declare Function ExtReceiverInputReadV2 Lib "r3cu_AUTO_200.DLL" ( ByVal FlagInput As Integer , ByVal iInputId As Integer , ByRef iValue As Integer ) As Integer 
Declare Function ExtReceiverOutputWrite Lib "r3cu_AUTO_200.DLL" ( ByVal iOutputId As Integer , ByVal iValue As Integer ) As Integer 
Declare Function AutomationOutputWrite Lib "r3cu_AUTO_200.DLL" ( ByVal iOutputId As Integer , ByVal iValue As Integer ) As Integer 
Declare Function AutomationGetHomingNeeded Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Sub AutomationSetHomingNeeded Lib "r3cu_AUTO_200.DLL" ( ByVal HomingNeeded As Integer )   
Declare Function AutomationGetHomingResult Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Sub AutomationSetHomingResult Lib "r3cu_AUTO_200.DLL" ( ByVal HomingOk As Integer )   
'  BOOL WINAPI AutomationGetStatus (typAUTOSTATUS *Status, long Head);
'  BOOL WINAPI AutomationSetStatus (typAUTOSTATUS *Status, long Head);
'  BOOL WINAPI AutomationGetPcStatus (typAUTOPCSTATUS *Status, long Head);
'  BOOL WINAPI AutomationSetPcStatus (typAUTOPCSTATUS *Status, long Head);
'  BOOL WINAPI AutomationSetSysCfg (typAUTOSYSCFG *SysCfg);
'  BOOL WINAPI AutomationSetTprjCfg (typAUTOTPRJCFG *TprjCfg, BOOL SystemFlag);
'  BOOL WINAPI AutomationSetRampTprjCfg (typAUTORAMPTPRJCFG *RampTprjCfg);
Declare Function IsReceiverDown Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function IsReceiverMiddle Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function IsReceiverUp Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationSetCurrentHead Lib "r3cu_AUTO_200.DLL" ( ByVal Head As Integer ) As Integer 
'  BOOL WINAPI AutomationGetSysCfg (typAUTOSYSCFG *SysCfg);
'  BOOL WINAPI AutomationGetTprjCfg (typAUTOTPRJCFG *TprjCfg, BOOL SystemFlag);
'  BOOL WINAPI AutomationGetRampTprjCfg (typAUTORAMPTPRJCFG *RampTprjCfg);
Declare Function AutomationGetPosition Lib "r3cu_AUTO_200.DLL" ( ByRef Position As Integer , ByVal Head As Integer ) As Integer 
Declare Function AutomationSetPosition Lib "r3cu_AUTO_200.DLL" ( ByVal Position As Integer , ByVal Head As Integer ) As Integer 
Declare Function AutomationSetMovementPending Lib "r3cu_AUTO_200.DLL" ( ByVal MoveType As Integer , ByVal Head As Integer ) As Integer 
Declare Function AutomationGetMovementPending Lib "r3cu_AUTO_200.DLL" ( ByRef MoveType As Integer , ByVal Head As Integer ) As Integer 
Declare Function AutomationSetCmdReceived Lib "r3cu_AUTO_200.DLL" ( ByVal CmdCode As Integer , ByVal Head As Integer ) As Integer 
Declare Function AutomationGetCmdReceived Lib "r3cu_AUTO_200.DLL" ( ByRef CmdCode As Integer , ByRef Head As Integer ) As Integer 
Declare Function ConveyorGetPosition Lib "r3cu_AUTO_200.DLL" ( ByRef Position As Integer , ByVal Head As Integer ) As Integer 
Declare Function ConveyorSetPosition Lib "r3cu_AUTO_200.DLL" ( ByVal Position As Integer , ByVal Head As Integer ) As Integer 
Declare Function ConveyorGetHomingNeeded Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Sub ConveyorSetHomingNeeded Lib "r3cu_AUTO_200.DLL" ( ByVal HomingNeeded As Integer )   
Declare Function ConveyorGetHomingResult Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Sub ConveyorSetHomingResult Lib "r3cu_AUTO_200.DLL" ( ByVal HomingOk As Integer )   
Declare Function ConveyorSetStepMoving Lib "r3cu_AUTO_200.DLL" ( ByVal MovingFlag As Integer , ByVal Head As Integer ) As Integer 
Declare Function ConveyorGetStepMoving Lib "r3cu_AUTO_200.DLL" ( ByRef MovingFlag As Integer , ByVal Head As Integer ) As Integer 
Declare Function ShifterSetStepMoving Lib "r3cu_AUTO_200.DLL" ( ByVal MovingFlag As Integer , ByVal Head As Integer ) As Integer 
Declare Function ShifterGetStepMoving Lib "r3cu_AUTO_200.DLL" ( ByRef MovingFlag As Integer , ByVal Head As Integer ) As Integer 
Declare Function AutomationExtSmemaIn Lib "r3cu_AUTO_200.DLL" ( ByVal LineWidth As Integer ) As Integer 
Declare Function AutomationExtSmemaOut Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationGetWizardRunning Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Sub AutomationSetWizardRunning Lib "r3cu_AUTO_200.DLL" ( ByVal WizardRunning As Integer )   
'  BOOL WINAPI AutomationGetDetailedStatus (typAUTODETSTATUS *Status, long Head);
'  BOOL WINAPI AutomationSetDetailedStatus (typAUTODETSTATUS *Status, long Head);
'  BOOL WINAPI AutomationGetStatusEx (typAUTOSTATUSEX *Status, long Head);
'  BOOL WINAPI AutomationSetStatusEx (typAUTOSTATUSEX *Status, long Head);
'  BOOL WINAPI AutomationGetPcStatusEx (typAUTOPCSTATUSEX *Status, long Head);
'  BOOL WINAPI AutomationSetPcStatusEx (typAUTOPCSTATUSEX *Status, long Head);
Declare Function AutomationGetFlyStatus Lib "r3cu_AUTO_200.DLL" ( ByVal pPos As Integer , ByRef pStatus As Integer , ByVal pHead As Integer ) As Integer 
Declare Function AutomationSetFlyStatus Lib "r3cu_AUTO_200.DLL" ( ByVal pPos As Integer , ByRef Status As Integer , ByVal pHead As Integer ) As Integer 
'  BOOL WINAPI AutomationGetPcStatus2 (typAUTOPCSTATUS2 *pStatus, long pHead);
'  BOOL WINAPI AutomationSetPcStatus2 (typAUTOPCSTATUS2 *Status, long pHead);
'  BOOL WINAPI AutomationGetLightStatus (typAUTOLIGHTSTATUS *pLightStatus, long pHead);
'  BOOL WINAPI AutomationSetLightStatus (typAUTOLIGHTSTATUS *pLightStatus, long pHead);
'  BOOL WINAPI AutomationGetTestCellPrevSmemaResult (UCHAR *pPrevSmemaResult, long pHead);
'  BOOL WINAPI AutomationSetTestCellPrevSmemaResult (UCHAR pPrevSmemaResult, long pHead);
'  BOOL WINAPI AutomationGetTestCellCurrSmemaResult (UCHAR *pCurrSmemaResult, long pHead);
'  BOOL WINAPI AutomationSetTestCellCurrSmemaResult (UCHAR pCurrSmemaResult, long pHead);
'  BOOL WINAPI AutomationGetSwModeBooking (UCHAR *pSwModeBooking);
'  BOOL WINAPI AutomationSetSwModeBooking (UCHAR pSwModeBooking);
Declare Function AutomationBoardMarkerSet Lib "r3cu_AUTO_200.DLL" ( ByVal pMarkerTime As Integer ) As Integer 
Declare Function AutomationBoardMarkerUpSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationConveyorEmptyCheck Lib "r3cu_AUTO_200.DLL" ( ByRef pFlagEmptyError As Integer ) As Integer 
Declare Function AutomationGetReadableCode Lib "r3cu_AUTO_200.DLL" ( ByVal pCodeId As Integer , ByRef pCodeData As Integer ) As Integer 
Declare Function AutomationSetReadableCode Lib "r3cu_AUTO_200.DLL" ( ByVal pCodeId As Integer , ByVal pCodeData As Integer ) As Integer 
'  BOOL WINAPI AutomationGetStatusFly (typAUTOSTATUS_FLY *pStatusFly, long pHead);
Declare Function AutomationGetSensorWizardStatus Lib "r3cu_AUTO_200.DLL" ( ByRef pStatus As Integer , ByVal pHead As Integer ) As Integer 
Declare Function AutomationSetSensorWizardStatus Lib "r3cu_AUTO_200.DLL" ( ByRef pStatus As Integer , ByVal pHead As Integer ) As Integer 
'  BOOL WINAPI AutomationGetPowerSaveStatus (UCHAR *pPowerSaveStatus, long pHead);
'  BOOL WINAPI AutomationSetPowerSaveStatus (UCHAR pPowerSaveStatus, long pHead);
Declare Function AutomationExtendedAreaPositionSet Lib "r3cu_AUTO_200.DLL" ( ByVal pEnableFlag As Integer , ByVal pExtendedAreaPositionId As Integer ) As Integer 
Declare Function AutomationOverturnWizardClose Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function AutomationInLineProductChange Lib "r3cu_AUTO_200.DLL" ( ByVal pPhase As Integer , ByVal pResult As Integer ) As Integer 
'  BOOL WINAPI AutomationGetAckInLineProductChange (UCHAR *pAckInLineProductChangeReceived, UCHAR *pAckInLineProductChange, long pHead);
'  BOOL WINAPI AutomationSetAckInLineProductChange (UCHAR pAckInLineProductChangeReceived, UCHAR pAckInLineProductChange, long pHead);
Declare Function AutomationBoardReposition Lib "r3cu_AUTO_200.DLL" ( ByVal pPosition As Integer ) As Integer 
Declare Function AutomationIsNewEddris Lib "r3cu_AUTO_200.DLL" ( ByRef pIsNewEddris As Integer , ByVal pElectricLevel As String ) As Integer 
'  BOOL WINAPI AutomationPassCheckEnable (UCHAR pFlag);
'  BOOL WINAPI AutomationPassCheckResultSet (UCHAR pResult);
Declare Function AutomationPassCheckLineStatusGet Lib "r3cu_AUTO_200.DLL" ( ByRef pConveyorInStatus As Integer , ByRef pStage1Status As Integer , ByRef pStage2Status As Integer , ByRef pConveyorOutStatus As Integer ) As Integer 
Declare Function AutomationWaitCriticalSection Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
'  BOOL WINAPI AutomationSkipTestEnable (UCHAR pFlag);
Declare Function ClampLock Lib "r3cu_AUTO_200.DLL" ( ByVal pFront As Integer , ByVal pRear As Integer ) As Integer 
Declare Function ClampUnlock Lib "r3cu_AUTO_200.DLL" ( ByVal pFront As Integer , ByVal pRear As Integer ) As Integer 
Declare Sub ReceiverDown Lib "r3cu_AUTO_200.DLL" ( )   
Declare Sub ReceiverMiddle Lib "r3cu_AUTO_200.DLL" ( )   
Declare Sub ReceiverUp Lib "r3cu_AUTO_200.DLL" ( )   
Declare Sub ReceiverMoveTo Lib "r3cu_AUTO_200.DLL" ( ByVal pPos As Double )   
Declare Function IsSystemManual Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Sub ConveyorNextStep Lib "r3cu_AUTO_200.DLL" ( )   
Declare Sub ConveyorStepMove Lib "r3cu_AUTO_200.DLL" ( ByVal StepQty As Integer )   
Declare Function IsConveyorAtPosition Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function DrawerLock Lib "r3cu_AUTO_200.DLL" ( ByVal LockFlag As Integer ) As Integer 
Declare Function ShuttleSystemSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ShuttleTprjSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ConveyorSystemSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function ConveyorTprjSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Sub ShifterHoming Lib "r3cu_AUTO_200.DLL" ( )   
Declare Sub ShifterPositionSet Lib "r3cu_AUTO_200.DLL" ( ByVal Pos As Integer )   
Declare Function IsShifterAtPosition Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function IsFixtureLocked Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function IsFixturePresent Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function IsPressurePlatePresent Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function LoaderSystemSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function UnloaderSystemSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function LoaderTprjSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
Declare Function UnloaderTprjSetupSet Lib "r3cu_AUTO_200.DLL" ( ) As Integer 
'if !defined(__AUTO_CONSTS_H)
Public Const uuAUTO_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_AUTO_UNIT As String =            "RP3_AutoUnit"
Public Const AUTOTYPE_NONE As Integer =    0
Public Const AUTOTYPE_CAN As Integer =     1
Public Const AUTOTYPE_SERMISC As Integer = 2
Public Const AUTOTYPE_RS232 As Integer =   3
Public Const AUTOTYPE_PARMISC As Integer = 4
Public Const AUTOTYPE_FLYPROBE As Integer =5
Public Const AUTO_SWSTATUS_NONE As Integer =  0
Public Const AUTO_SWSTATUS_READY As Integer = 1
Public Const AUTO_SWSTATUS_RUN As Integer =   2
Public Const AUTO_SWSTATUS_BUSY As Integer =  3
Public Const AUTO_SWSTATUS_PASS As Integer =  4
Public Const AUTO_SWSTATUS_FAIL As Integer =  5
Public Const AUTO_SWSTATUS_ALARM As Integer =   6
Public Const AUTO_SWSTATUS_NOALARM As Integer = 7
Public Const AUTO_MODE_NONE As Integer =    &HFF
Public Const AUTO_MODE_MANUAL As Integer =  0
Public Const AUTO_MODE_LOCAL As Integer =   1
Public Const AUTO_MODE_INLINE As Integer =  2
Public Const AUTO_MODE_MASK As Integer =    &H0F
Public Const AUTO_MODE_BOOKING As Integer = &H10
Public Const AUTO_MAX_HEAD As Integer =  4
Public Const AUTO_CURRENT_HEAD As Integer = &HF0
Public Const AUTO_ALL_HEADS As Integer =    &HFF
Public Const AUTO_BUZZER_OFF As Integer =  0
Public Const AUTO_BUZZER_SLOW As Integer = 1
Public Const AUTO_BUZZER_FAST As Integer = 2
Public Const AUTO_SEMAPHORE_GREEN As Integer =  0
Public Const AUTO_SEMAPHORE_ORANGE As Integer = 1
Public Const AUTO_SEMAPHORE_RED As Integer =    2
Public Const AUTO_SEMAPHORE_BLUE As Integer =   3
Public Const AUTO_CMD_NONE As Integer =                 &H00    ' No command
Public Const AUTO_CMD_CONSOLE_STOP As Integer =         &H01    ' Stop from conso&e
Public Const AUTO_CMD_CONSOLE_START As Integer =        &H02    ' Start from conso&e
Public Const AUTO_CMD_CONSOLE_SKIP As Integer =         &H03    ' Skip from conso&e
Public Const AUTO_CMD_STATUS_UPDATED As Integer =       &H04    ' Status updated
Public Const AUTO_CMD_CONSOLE_RESERV As Integer =       &H05    ' Reserve Start from conso&e
Public Const AUTO_CMD_CONSOLE_STOP_ERR As Integer =     &H07    ' Stop from conso&e per errore misc (no stop misc)
Public Const AUTO_CMD_CONSOLE_CONTINUE As Integer =     &H08    ' Continue
Public Const AUTO_CMD_CONSOLE_START_HOLD As Integer =   &H09    ' Start from conso&e for ho&d cap
Public Const AUTO_CMD_CONSOLE_RESERV_RESET As Integer = &H0A    ' Reserve Start from conso&e reset
Public Const AUTO_CMD_NOT_HANDLED_MSG As Integer =      &H7F    ' Not managed message
Public Const AUTO_DIRECTION_LEFT As Integer =   0
Public Const AUTO_DIRECTION_RIGHT As Integer =  1
Public Const AUTO_SWERR_NONE As Integer =   0
Public Const AUTO_SWERR_HIGH As Integer =   1
Public Const AUTO_SWERR_MEDIUM As Integer = 2
Public Const AUTO_SWERR_LOW As Integer =    3
Public Const AUTO_SWERR_WRN_FAIL As Integer = 4    ' Warning on repetitive fails
Public Const AUTO_FAILMODE_PASSTH As Integer =   0
Public Const AUTO_FAILMODE_PASSBACK As Integer = 1
Public Const AUTO_FAILMODE_BLOCK As Integer =    2
Public Const AUTO_SNFAILMODE_LEGACY As Integer =   &H00
Public Const AUTO_SNFAILMODE_SKIP As Integer =     &H10
Public Const AUTO_SNFAILMODE_BLOCK As Integer =    &H20
Public Const AUTO_SNFAILMODE_START As Integer =    &H30
Public Const AUTO_SN_NONE As Integer = 0
Public Const AUTO_SN_OK As Integer =   1
Public Const AUTO_SN_KO As Integer =   2
Public Const AUTO_LOCK_SINGLE As Integer =          1
Public Const AUTO_LOCK_SINGLE_ENCODER As Integer =  2
Public Const AUTO_LOCK_DUAL As Integer =            3
Public Const AUTO_LOCK_DUAL_ENCODER As Integer =    4
Public Const AUTO_STOPPER_ELECTRIC As Integer =     1
Public Const AUTO_STOPPER_PNEUMATIC As Integer =    2
Public Const AUTO_STOPPER_NOT_PRESENT As Integer =  3
Public Const AUTO_STOPPER_DEFAULT_POS As Integer =  &H00
Public Const AUTO_STOPPER_LEFT_POS As Integer =     &H01
Public Const AUTO_STOPPER_RIGHT_POS As Integer =    &H02
Public Const AUTO_BRCSETUP_ON_SENSOR As Integer =  &H01   ' &ettura barcode da fermo su& sensore
Public Const AUTO_BRCSETUP_DEFAULT As Integer =    &H00
Public Const AUTO_CONVEY_STANDARD As Integer = 0
Public Const AUTO_CONVEY_450 As Integer =      1
Public Const AUTO_MWA_450 As Integer =         2
Public Const AUTO_CONVEY_450_MWA As Integer =  3
Public Const AUTO_SLIM As Integer =            4
Public Const AUTO_PRESSOR_IRM100 As Integer = 0
Public Const AUTO_PRESSOR_IRM200 As Integer = 1
Public Const AUTO_PRESSOR_IRM210 As Integer = 2
Public Const AUTO_DOOR_PRESENT As Integer = 0
Public Const AUTO_DOOR_OPTBARR As Integer = 1
Public Const AUTO_DOOR_ABSENT As Integer =  2
Public Const ACKSTACK_INVALID As Integer = &HFF
Public Const ACKSTACK_SIZE As Integer = 100
Public Const AUTO_ROTFAILMODE_BLOCK As Integer =  0
Public Const AUTO_ROTFAILMODE_SKIP As Integer =   1
Public Const AUTO_BRDSKIP_SNMISSING As Integer =  0
Public Const AUTO_LIGHTINT_LOW As Integer =    0
Public Const AUTO_LIGHTINT_MEDIUM As Integer = 1
Public Const AUTO_LIGHTINT_HIGH As Integer =   2
Public Const AUTO_MRXX_HOMING As Integer = -35
Public Const AUTO_IRMS_HOMING As Integer = -35
Public Const AUTO_ERRMODE_SKIP As Integer =   0
Public Const AUTO_ERRMODE_BLOCK As Integer =  1
Public Const AUTO_CONVEYOR_WIDTH_MAX As Integer = -1111
Public Const AUTO_CONVEYOR_WIDTH_RSL As Integer = -2222
Public Const AUTO_BTU_ORIENTATION_EMPTY As Integer = &HFF
Public Const AUTO_BTU_ORIENTATION_0 As Integer =     0
Public Const AUTO_BTU_ORIENTATION_90 As Integer =    1
Public Const AUTO_AXIS_DOOR_MANUAL As Integer =             &H00
Public Const AUTO_AXIS_DOOR_MANUAL_WITH_KEY As Integer =    &H01
Public Const AUTO_AXIS_DOOR_SLIDING_MANUAL As Integer =     &H02
Public Const AUTO_AXIS_DOOR_SLIDING_MOTORIZED As Integer =  &H03
'endif
'if !defined(__LCS_H)
Public Const uuLCS_H As Integer = 0
'include "r3cuLcsConsts.h"
'ifdef __cplusplus
Declare Function fLcsLibraryAttach Lib "R3CU_LCS_200.DLL" ( ByVal pAllocData As Integer ) As Integer 
Declare Function fLcsLibraryDetach Lib "R3CU_LCS_200.DLL" ( ) As Integer 
Declare Function LcsEepromAnagraphWrite Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pDataBuffer As String , ByVal pDataSize As Integer , ByVal pOffset As Integer ) As Integer 
Declare Function LcsEepromAnagraphRead Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pDataBuffer As String , ByVal pDataSize As Integer , ByVal pOffset As Integer ) As Integer 
Declare Function LcsEepromCalibrationWrite Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pDataBuffer As String , ByVal pDataSize As Integer , ByVal pOffset As Integer ) As Integer 
Declare Function LcsEepromCalibrationRead Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pDataBuffer As String , ByVal pDataSize As Integer , ByVal pOffset As Integer ) As Integer 
Declare Function LcsSetupSet Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pTime As Integer , ByVal pColorGain As Integer , ByVal pLuxGain As Integer ) As Integer 
Declare Function LcsSetupAutoscale Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pTime As Integer , ByRef pColorGain As Integer , ByRef pLuxGain As Integer ) As Integer 
Declare Function LcsSingleTestExecute Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pCalibrationUsed As Integer , ByRef pTestResult As Integer ) As Integer 
Declare Function LcsParallelTestExecute Lib "R3CU_LCS_200.DLL" ( ByRef pSensorIdList As Integer , ByVal pCalibrationUsed As Integer , ByRef pTestResultList As Integer ) As Integer 
Declare Function LcsMeasuredValuesRead Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByRef pRgbRed As Double , ByRef pRgbGreen As Double , ByRef pRgbBlue As Double , ByRef pHslHue As Double , ByRef pHslSaturation As Double , ByRef pHslLightness As Double , ByRef pXyX As Double , ByRef pXyY As Double , ByRef pLuxVisible As Double , ByRef pLuxInfrared As Double ) As Integer 
Declare Function LcsSensorIdWrite Lib "R3CU_LCS_200.DLL" ( ByVal pSensorSerialNumber As String , ByVal pSensorId As Integer ) As Integer 
Declare Function LcsSensorIdRead Lib "R3CU_LCS_200.DLL" ( ByVal pSensorSerialNumber As String , ByRef pSensorId As Integer ) As Integer 
Declare Function LcsSensorsScan Lib "R3CU_LCS_200.DLL" ( ByRef pSensorsFoundQty As Integer , ByRef pSensorIdList As Integer , ByVal pSensorSerialNumberList As String , ByRef pSensorPortList As Integer ) As Integer 
Declare Function LcsInputRead Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pAddress As Integer , ByRef pStatus As Integer ) As Integer 
Declare Function LcsOutputRead Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pAddress As Integer , ByRef pStatus As Integer ) As Integer 
Declare Function LcsOutputWrite Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pAddress As Integer , ByVal pStatus As Integer ) As Integer 
Declare Function LcsBusyRead Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByRef pBusyStatus As Integer ) As Integer 
Declare Function LcsTestResultRead Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByRef pTestResult As Integer ) As Integer 
Declare Function LcsTestStop Lib "R3CU_LCS_200.DLL" ( ) As Integer 
Declare Function LcsLedSet Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pRedStatus As Integer , ByVal pGreenStatus As Integer , ByVal pBlueStatus As Integer ) As Integer 
Declare Function LcsSensorFinderOn Lib "R3CU_LCS_200.DLL" ( ByVal pLuxLimit As Integer ) As Integer 
Declare Function LcsSensorFinderOff Lib "R3CU_LCS_200.DLL" ( ) As Integer 
Declare Function LcsSensorFinderRead Lib "R3CU_LCS_200.DLL" ( ByRef pSensorFound As Integer , ByVal pSensorSerialNumber As String , ByRef pSensorId As Integer ) As Integer 
Declare Function LcsThresholdSet Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pRedMin As Double , ByVal pRedMax As Double , ByVal pGreenMin As Double , ByVal pGreenMax As Double , ByVal pBlueMin As Double , ByVal pBlueMax As Double , ByVal pHueMin As Double , ByVal pHueMax As Double , ByVal pSaturationMin As Double , ByVal pSaturationMax As Double , ByVal pLightnessMin As Double , ByVal pLightnessMax As Double , ByVal pXMin As Double , ByVal pXMax As Double , ByVal pYMin As Double , ByVal pYMax As Double , ByVal pLuxMin As Double , ByVal pLuxMax As Double , ByVal pLuxIrMin As Double , ByVal pLuxIrMax As Double ) As Integer 
Declare Function LcsTrimSet Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pRed As Double , ByVal pGreen As Double , ByVal pBlue As Double , ByVal pHue As Double , ByVal pSaturation As Double , ByVal pLightness As Double , ByVal pX As Double , ByVal pY As Double , ByVal pLux As Double , ByVal pLuxIr As Double ) As Integer 
Declare Function LcsTrimEnabledSet Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pRed As Byte , ByVal pGreen As Byte , ByVal pBlue As Byte , ByVal pHue As Byte , ByVal pSaturation As Byte , ByVal pLightness As Byte , ByVal pX As Byte , ByVal pY As Byte , ByVal pLux As Byte , ByVal pLuxIr As Byte ) As Integer 
Declare Function LcsMeasEnabledSet Lib "R3CU_LCS_200.DLL" ( ByVal pSensorId As Integer , ByVal pRGB As Byte , ByVal pHSL As Byte , ByVal pXY As Byte , ByVal pLux As Byte , ByVal pLuxIr As Byte ) As Integer 
'if !defined(__LCS_CONSTS_H)
Public Const uuLCS_CONSTS_H As Integer = 0
Public Const kTIME_12_ms As Integer =                              100
Public Const kTIME_100_ms As Integer =                             101
Public Const kTIME_700_ms As Integer =                             102
Public Const kCOLOR_GAIN_1 As Integer =                            103
Public Const kCOLOR_GAIN_4 As Integer =                            104
Public Const kCOLOR_GAIN_16 As Integer =                           105
Public Const kCOLOR_GAIN_64 As Integer =                           106
Public Const kLUX_GAIN_1 As Integer =                              107
Public Const kLUX_GAIN_8 As Integer =                              108
Public Const kLUX_GAIN_16 As Integer =                             109
Public Const kLUX_GAIN_120 As Integer =                            110
Public Const kCOLOR_PRESCALER_1 As Integer =                       111
Public Const kCOLOR_PRESCALER_2 As Integer =                       112
Public Const kCOLOR_PRESCALER_4 As Integer =                       113
Public Const kCOLOR_PRESCALER_8 As Integer =                       114
Public Const kCOLOR_PRESCALER_16 As Integer =                      115
Public Const kCOLOR_PRESCALER_32 As Integer =                      116
Public Const kCOLOR_PRESCALER_64 As Integer =                      117
Public Const kLUX_PRESCALER_6_25 As Integer =                      118
Public Const kERR_LCS_TIMEOUT As Integer =                        -100
Public Const kERR_LCS_TIME As Integer =                           -101
Public Const kERR_LCS_COLOR_GAIN As Integer =                     -102
Public Const kERR_LCS_LUX_GAIN As Integer =                       -103
Public Const kERR_LCS_SENSOR_ID As Integer =                      -104
Public Const kERR_LCS_PORT As Integer =                           -105
Public Const kERR_LCS_NO_DATA_READ As Integer =                   -106
Public Const kERR_LCS_ID_NOT_VALID As Integer =                   -200
Public Const kERR_LCS_FW_TIMEOUT As Integer =                     -201
Public Const kERR_LCS_MEASURE_NOT_ENABLED As Integer =            -202
Public Const kERR_LCS_MEAS_STAT_ALS_NOT_READ As Integer =         -203
Public Const kERR_LCS_MEAS_STAT_ALS_READ As Integer =             -204
Public Const kERR_LCS_MEAS_STAT_ALS_DONE_GOOD As Integer =        -205
Public Const kERR_LCS_MEAS_STAT_ALS_DONE_UNDERRANGE As Integer =  -206
Public Const kERR_LCS_MEAS_STAT_ALS_DONE_OVERRANGE As Integer =   -207
Public Const kERR_LCS_MEAS_STAT_RGB_NOT_READ As Integer =         -208
Public Const kERR_LCS_MEAS_STAT_RGB_READ As Integer =             -209
Public Const kERR_LCS_MEAS_STAT_RGB_DONE_GOOD As Integer =        -210
Public Const kERR_LCS_MEAS_STAT_RGB_DONE_UNDERRANGE As Integer =  -211
Public Const kERR_LCS_MEAS_STAT_RGB_DONE_OVERRANGE As Integer =   -212
Public Const kERR_LCS_MATRIX_NOT_VALID As Integer =               -213
Public Const kERR_LCS_COMMUNICATION_RGBC As Integer =             -214
Public Const kERR_LCS_COMMUNICATION_ALS As Integer =              -215
Public Const kERR_LCS_COMMUNICATION_EEPROM As Integer =           -216
Public Const kERR_LCS_CALIB_READ As Integer =                     -217
Public Const kERR_LCS_CALIB_NOT_VALID As Integer =                -218
Public Const kERR_LCS_CALIB_NOT_PRESENT As Integer =              -219
Public Const kERR_LCS_CALIB_CHECKSUM_ERROR As Integer =           -220
Public Const kERR_LCS_EEPROM_PARAMETER_NOT_VALID As Integer =     -221
Public Const kERR_LCS_COMM1 As Integer =                          -222
Public Const kERR_LCS_COMM2 As Integer =                          -223
Public Const kERR_LCS_COMM3 As Integer =                          -224
Public Const kERR_LCS_COMM4 As Integer =                          -225
Public Const kERR_LCS_COMM5 As Integer =                          -226
Public Const kERR_LCS_COMM6 As Integer =                          -227
Public Const kERR_LCS_COMM7 As Integer =                          -228
Public Const kERR_LCS_PARAM1 As Integer =                         -229
Public Const kERR_LCS_PARAM2 As Integer =                         -230
Public Const kERR_LCS_PARAM3 As Integer =                         -231
Public Const kERR_LCS_PARAM4 As Integer =                         -232
Public Const kERR_LCS_PARAM5 As Integer =                         -233
'endif
Structure TTestMeas
  Dim TprjId As Integer
  Dim TplanId As Integer
  Dim Site As Integer
  <VBFixedString(32)> Dim TaskName As String
  Dim TaskNumber As Integer
  Dim TestNumber As Integer
  Dim UniqueTestId As Integer
  <VBFixedString(128)> Dim TestRemark As String
  Dim TestResult As Integer
  Dim MeasuredValue As Double
  Dim HighThreshold As Double
  Dim LowThreshold As Double
  <VBFixedString(8)> Dim Unit As String
End Structure

Declare Function ReadAnlTaskMeas Lib "r3cu_COMMON_200.DLL" ( ByVal TestplanName As String , ByVal StartTaskName As String , ByVal EndTaskName As String , TestMeasArray As TTestMeas, ByVal ArraySize As Integer , StoredMeasSize As Integer ) As Integer 
Declare Function ReadAnlTaskLabelMeas Lib "r3cu_COMMON_200.DLL" ( ByVal TestplanName As String , ByVal StartTaskLabel As String , ByVal EndTaskLabel As String , TestMeasArray As TTestMeas, ByVal ArraySize As Integer , StoredMeasSize As Integer ) As Integer 
Declare Function TplanAddControlFunctions Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal tplanFuncToAdd As Integer ) As Integer 
'if !defined(__PMU_H)
Public Const uuPMU_H As Integer = 0
'include "r3cuPmuConsts.h"
'include "r3cuPmuTypes.h"
'ifdef __cplusplus
Declare Function PmuLibraryAttach Lib "r3cu_PMU_200.DLL" ( ) As Integer 
Declare Function PmuLibraryDetach Lib "r3cu_PMU_200.DLL" ( ) As Integer 
Declare Sub PmuDefaultsSet Lib "r3cu_PMU_200.DLL" ( ByVal vbEvent As Integer )   
Declare Function DvmConnectInterf Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Hot As Integer , ByVal Cold As Integer ) As Integer 
Declare Function DvmDisconnectInterf Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Hot As Integer , ByVal Cold As Integer ) As Integer 
Declare Function DvmConnectAbus Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Hot As Integer , ByVal Cold As Integer ) As Integer 
Declare Function DvmDisconnectAbus Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Hot As Integer , ByVal Cold As Integer ) As Integer 
Declare Function DvmConnectMbus Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Hot As Integer , ByVal Cold As Integer ) As Integer 
Declare Function DvmDisconnectMbus Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Hot As Integer , ByVal Cold As Integer ) As Integer 
Declare Function DvmDisconnectAll Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
'  long WINAPI pDvmConnectVref (long InstrId);
'  long WINAPI pDvmDisconnectVref (long InstrId);
Declare Function CntConnectInterf Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Hot As Integer , ByVal Cold As Integer , ByVal Ref As Integer ) As Integer 
Declare Function CntDisconnectInterf Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Hot As Integer , ByVal Cold As Integer , ByVal Ref As Integer ) As Integer 
Declare Function CntConnectAbus Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal ChA As Integer , ByVal ChB As Integer , ByVal Ref As Integer ) As Integer 
Declare Function CntDisconnectAbus Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal ChA As Integer , ByVal ChB As Integer , ByVal Ref As Integer ) As Integer 
Declare Function CntConnectMbus Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal ChA As Integer , ByVal ChB As Integer , ByVal Ref As Integer ) As Integer 
Declare Function CntDisconnectMbus Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal ChA As Integer , ByVal ChB As Integer , ByVal Ref As Integer ) As Integer 
Declare Function CntConnectFlag Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal ChA As Integer , ByVal ChB As Integer ) As Integer 
Declare Function CntDisconnectFlag Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal ChA As Integer , ByVal ChB As Integer ) As Integer 
Declare Function CntDisconnectAll Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function CntConnectDvm Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal ChA As Integer , ByVal ChB As Integer ) As Integer 
Declare Function CntDisconnectDvm Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal ChA As Integer , ByVal ChB As Integer ) As Integer 
Declare Function DvmMbusReadConversionEnable Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DvmMbusReadConversionDisable Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DvmClear Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DvmSet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal InputStage As Integer , ByVal Coupling As Integer , ByVal Filter As Integer , ByVal VRange As Integer , ByVal MeasMode As Integer , ByVal MeasType As Integer , ByVal EnableAcqRam As Integer ) As Integer 
Declare Function DvmTrigSet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal iInput As Integer , ByVal Slope As Integer , ByVal Vtrigger As Double ) As Integer 
Declare Function DvmSyncSet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal FlagId As Integer , ByVal SbusLine As Integer ) As Integer 
Declare Function DvmSyncPolaritySet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal FlagId As Integer , ByVal Polarity As Integer ) As Integer 
Declare Function DvmCompSet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal VthLow As Double , ByVal VthHigh As Double ) As Integer 
Declare Function DvmEnable Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function DvmDisable Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function CntClear Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function CntFreqSet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Resol As Integer , ByVal Range As Integer , ByVal Threshold As Double , ByVal Slope As Integer ) As Integer 
Declare Function CntPeriodSet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Resol As Integer , ByVal Range As Integer , ByVal Threshold As Double , ByVal Slope As Integer ) As Integer 
Declare Function CntIntervalSet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Resol As Integer , ByVal ARange As Integer , ByVal AThreshold As Double , ByVal ASlope As Integer , ByVal BRange As Integer , ByVal BThreshold As Double , ByVal BSlope As Integer ) As Integer 
Declare Function CntRatioSet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Resol As Integer , ByVal ARange As Integer , ByVal AThreshold As Double , ByVal ASlope As Integer , ByVal BRange As Integer , ByVal BThreshold As Double , ByVal BSlope As Integer ) As Integer 
Declare Function CntEventSet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Range As Integer , ByVal Threshold As Double , ByVal Slope As Integer ) As Integer 
Declare Function CntSyncSet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal FlagId As Integer , ByVal SbusLine As Integer ) As Integer 
Declare Function CntSyncPolaritySet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal FlagId As Integer , ByVal Polarity As Integer ) As Integer 
Declare Function CntChannelPolaritySet Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal A_Polarity As Integer , ByVal B_Polarity As Integer ) As Integer 
Declare Function CntEnable Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function CntDisable Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function CntTimeOut Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal Time As Integer ) As Integer 
Declare Function DvmRead Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByRef Value As Double , ByRef MeasType As Integer ) As Integer 
Declare Function DvmReadWindows Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByRef Reasult As Integer ) As Integer 
Declare Function DvmReadAll Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByRef ValueArray As Double , ByVal ValueArraySize As Integer , ByRef MeasType As Integer , ByRef MeasNum As Integer ) As Integer 
Declare Function DvmReadSignalChar Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByRef ValueMin As Double , ByRef ValueMax As Double , ByRef ValueAvg As Double , ByRef MeasType As Integer , ByRef MeasNum As Integer ) As Integer 
Declare Function CntRead Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByRef Value As Double , ByRef MeasType As Integer ) As Integer 
Declare Function DvmConnectUabusBmu Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal RowHot As Integer , ByVal RowCold As Integer ) As Integer 
Declare Function DvmDisconnectUabusBmu Lib "r3cu_PMU_200.DLL" ( ByVal InstrId As Integer , ByVal vbPoint As Integer ) As Integer 
Declare Sub DvmCmpSet Lib "r3cu_PMU_200.DLL" ( ByVal Mode As Integer , ByVal High As Double , ByVal Low As Double )   
Declare Sub DvmCompareSet Lib "r3cu_PMU_200.DLL" ( ByVal Mode As Integer , ByVal High As Double , ByVal Low As Double )   
Declare Sub DvmHandlingSet Lib "r3cu_PMU_200.DLL" ( ByVal InputStage As Integer , ByVal Coupling As Integer , ByVal Gain As Integer , ByVal Filter As Integer , ByVal Average As Integer )   
Declare Sub DvmConnectExtern Lib "r3cu_PMU_200.DLL" ( ByVal Via As Integer )   
Declare Sub DvmConnectNotStd Lib "r3cu_PMU_200.DLL" ( ByVal vbPoint As Integer , ByVal Row As Integer , ByVal Via As Integer )   
Declare Sub DvmInternalMeasureSet Lib "r3cu_PMU_200.DLL" ( ByVal UnitId As Integer , ByVal vbType As Integer , ByVal vbScale As Integer )   
Declare Sub DvmMbusSet Lib "r3cu_PMU_200.DLL" ( ByVal UnitId As Integer , ByVal vbType As Integer , ByVal vbScale As Integer )   
Declare Sub CntCounterSet Lib "r3cu_PMU_200.DLL" ( ByVal Slope As Integer , ByVal Trigger As Double )   
Declare Sub CntSet Lib "r3cu_PMU_200.DLL" ( ByVal InputStage As Integer , ByVal Coupling As Integer , ByVal Gain As Integer , ByVal Filter As Integer , ByVal Average As Integer )   
Declare Sub CntHandlingSet Lib "r3cu_PMU_200.DLL" ( ByVal InputStage As Integer , ByVal Coupling As Integer , ByVal Gain As Integer , ByVal Filter As Integer , ByVal Average As Integer )   
Declare Sub CntInternalMeasureSet Lib "r3cu_PMU_200.DLL" ( ByVal UnitId As Integer , ByVal vbType As Integer , ByVal vbScale As Integer )   
Declare Sub CntMbusSet Lib "r3cu_PMU_200.DLL" ( ByVal UnitId As Integer , ByVal vbType As Integer , ByVal vbScale As Integer )   
Declare Sub CntConnectExtern Lib "r3cu_PMU_200.DLL" ( ByVal Via As Integer )   
Declare Sub CntConnectNotStd Lib "r3cu_PMU_200.DLL" ( ByVal vbPoint As Integer , ByVal Row As Integer , ByVal Via As Integer )   
'if !defined(__PMU_CONSTS_H)
Public Const uuPMU_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_PMU_UNIT As String =            "RP3_PmuUnit"
Public Const UNKNOWN_PMU_RESULT As Integer =   0
'endif
'if !defined(__GEN_H)
Public Const uuGEN_H As Integer = 0
'ifdef __cplusplus
Declare Function WfgLibraryAttach Lib "r3cu_GEN_200.DLL" ( ) As Integer 
Declare Function WfgLibraryDetach Lib "r3cu_GEN_200.DLL" ( ) As Integer 
Declare Function WfgConnectAbus Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer , ByVal Row1 As Integer , ByVal Row2 As Integer ) As Integer 
Declare Function WfgConnectInterf Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function WfgConnectMobus Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer , ByVal Row As Integer ) As Integer 
Declare Function WfgDisconnectAbus Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer , ByVal Row1 As Integer , ByVal Row2 As Integer ) As Integer 
Declare Function WfgDisconnectInterf Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function WfgDisconnectMobus Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer , ByVal Row As Integer ) As Integer 
Declare Function WfgDisconnectAll Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function WfgInputSelect Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer , ByVal Source As Integer ) As Integer 
Declare Function WfgDisable Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function WfgEnable Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer ) As Integer 
Declare Function WfgOutputSet Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer , ByVal OutputMode As Integer , ByVal AmpRange As Integer , ByVal Amplitude As Double , ByVal OffRange As Integer , ByVal DCOffset As Double , ByVal Impedance As Integer ) As Integer 
Declare Function WfgTimingSet Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer , ByVal Flag As Integer ) As Integer 
Declare Function WfgWaveformSelect Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer , ByVal WformType As Integer , ByVal Frequency As Double , ByVal OutFormat As Integer ) As Integer 
Declare Function WfgSequenceSelect Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer , ByVal SequenceName As String , ByVal StepTime As Double , ByVal OutFormat As Integer ) As Integer 
Declare Function WfgConnectUabusBmu Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer , ByVal RowHot As Integer , ByVal RowCold As Integer ) As Integer 
Declare Function WfgDisconnectUabusBmu Lib "r3cu_GEN_200.DLL" ( ByVal InstrId As Integer ) As Integer 
'if !defined(__GEN_CONSTS_H)
Public Const uuGEN_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_GEN_UNIT As String =            "RP3_GenUnit"
'endif'if !defined(__AUX_H)
Public Const uuAUX_H As Integer = 0
'include "r3cuAuxConsts.h"
'include "r3cuAuxTypes.h"
'ifdef __cplusplus
Declare Function AuxLibraryAttach Lib "r3cu_AUX_200.DLL" ( ) As Integer 
Declare Function AuxLibraryDetach Lib "r3cu_AUX_200.DLL" ( ) As Integer 
Declare Function AuxGetSyncRegFromFlagId Lib "r3cu_AUX_200.DLL" ( ByVal FlagId As Integer ) As Integer 
'  'ifdef __cplusplus
'    long WINAPI AuxGetFlagIdFromSyncReg (long SyncReg, char *StrFlagId = NULL);
Declare Function AuxGetFlagIdFromSyncReg Lib "r3cu_AUX_200.DLL" ( ByVal SyncReg As Integer , ByVal StrFlagId As String ) As Integer 
'  'ifdef __cplusplus
'    long WINAPI SysGetSyncLineId (long SyncLine, char *StrSyncLine = NULL);
Declare Function SysGetSyncLineId Lib "r3cu_AUX_200.DLL" ( ByVal SyncLine As Integer , ByVal StrSyncLine As String ) As Integer 
Declare Function SysGetSyncRegFromFlagId Lib "r3cu_AUX_200.DLL" ( ByVal BoardId As Integer , ByVal FlagId As Integer ) As Integer 
'  'ifdef __cplusplus
'    long WINAPI SysGetFlagIdFromSyncReg (long BoardId, long SyncReg, char *StrFlagId = NULL);
Declare Function SysGetFlagIdFromSyncReg Lib "r3cu_AUX_200.DLL" ( ByVal BoardId As Integer , ByVal SyncReg As Integer , ByVal StrFlagId As String ) As Integer 
Declare Function SysConnectSyncLine Lib "r3cu_AUX_200.DLL" ( ByVal BoardId As Integer , ByVal BoardNo As Integer , ByVal FlagId As Integer , ByVal SyncLine As Integer ) As Integer 
Declare Function SysIdle Lib "r3cu_AUX_200.DLL" ( ) As Integer 
Declare Function AnlPhaseSet Lib "r3cu_AUX_200.DLL" ( ByVal PhaseId As Integer , ByVal Toff1 As Double , ByVal Ton As Double , ByVal Toff2 As Double , ByVal Cycles As Integer ) As Integer 
Declare Function AnlPhaseEnable Lib "r3cu_AUX_200.DLL" ( ByVal PhaseId As Integer ) As Integer 
Declare Function AnlPhaseDisable Lib "r3cu_AUX_200.DLL" ( ByVal PhaseId As Integer ) As Integer 
Declare Function AnlTimingEnable Lib "r3cu_AUX_200.DLL" ( ) As Integer 
Declare Function AnlTimingDisable Lib "r3cu_AUX_200.DLL" ( ) As Integer 
Declare Function AnlTimingOutSourceSet Lib "r3cu_AUX_200.DLL" ( ByVal OutSource As Integer ) As Integer 
Declare Function AnlTimingStartModeSet Lib "r3cu_AUX_200.DLL" ( ByVal StartSource As Integer , ByVal StartMode As Integer , ByVal StartPole As Integer ) As Integer 
Declare Function AnlTimeoutSet Lib "r3cu_AUX_200.DLL" ( ByVal TimeOut As Integer ) As Integer 
Declare Function AnlTimingSyncSet Lib "r3cu_AUX_200.DLL" ( ByVal FlagId As Integer , ByVal SbusLine As Integer ) As Integer 
Declare Function AnlTimingIdle Lib "r3cu_AUX_200.DLL" ( ) As Integer 
Declare Sub AnlTimingClear Lib "r3cu_AUX_200.DLL" ( )   
Declare Sub AnlTimingSet Lib "r3cu_AUX_200.DLL" ( ByVal Phase As Integer , ByVal TOn As Integer , ByVal TOff1 As Integer , ByVal TOff2 As Integer , ByVal Cycles As Integer )   
Declare Sub CntMeasTimOptimize Lib "r3cu_AUX_200.DLL" ( ByVal Enable As Integer )   
Declare Sub RelayTimeSet Lib "r3cu_AUX_200.DLL" ( ByVal WaitRelay As Integer )   
Declare Sub RelayModeSet Lib "r3cu_AUX_200.DLL" ( ByVal Mode As Integer )   
Declare Sub HvBreakerSet Lib "r3cu_AUX_200.DLL" ( ByVal Mode As Integer )   
Declare Function UserFlagsIdle Lib "r3cu_AUX_200.DLL" ( ) As Integer 
Declare Function UserFlagsSet Lib "r3cu_AUX_200.DLL" ( ByVal FlagGroup As Integer , ByRef FlagsArray As Integer ) As Integer 
Declare Function UserFlagsReset Lib "r3cu_AUX_200.DLL" ( ByVal FlagGroup As Integer , ByRef FlagsArray As Integer ) As Integer 
Declare Function UserFlagRead Lib "r3cu_AUX_200.DLL" ( ByVal FlagGroup As Integer , ByVal FlagId As Integer , ByRef FlagValue As Integer ) As Integer 
Declare Function UserFlagGroupRead Lib "r3cu_AUX_200.DLL" ( ByVal FlagGroup As Integer , ByRef FlagValuesArray As Integer ) As Integer 
Declare Function SysSyncSet Lib "r3cu_AUX_200.DLL" ( ByVal FlagId As Integer , ByVal SbusLine As Integer ) As Integer 
Declare Function ACLineSyncSet Lib "r3cu_AUX_200.DLL" ( ByVal FlagId As Integer , ByVal SbusLine As Integer ) As Integer 
Declare Function SysSyncRead Lib "r3cu_AUX_200.DLL" ( ByVal FlagId As Integer , ByRef vbData As Integer ) As Integer 
Declare Function SysSyncWrite Lib "r3cu_AUX_200.DLL" ( ByVal FlagId As Integer , ByVal vbData As Integer ) As Integer 
Declare Function EsbusLinkSbus Lib "r3cu_AUX_200.DLL" ( ByVal Mode As Integer , ByVal EsbusLine As Integer , ByVal SbusLine As Integer ) As Integer 
Declare Function SbusIdle Lib "r3cu_AUX_200.DLL" ( ) As Integer 
Declare Function SbusDefault Lib "r3cu_AUX_200.DLL" ( ByVal BusId As Integer ) As Integer 
Declare Function InterfSyncSet Lib "r3cu_AUX_200.DLL" ( ByVal FlagId As Integer , ByVal SbusLine As Integer ) As Integer 
Declare Function InterfConnectMobus Lib "r3cu_AUX_200.DLL" ( ByVal Modin As Integer ) As Integer 
Declare Function InterfDisconnectMobus Lib "r3cu_AUX_200.DLL" ( ByVal Modin As Integer ) As Integer 
Declare Function AnlRelayModeSet Lib "r3cu_AUX_200.DLL" ( ByVal RelayMode As Integer ) As Integer 
Declare Function AnlRelayTimeSet Lib "r3cu_AUX_200.DLL" ( ByVal RelayTime As Integer ) As Integer 
Declare Function RunA Lib "r3cu_AUX_200.DLL" ( ByVal TestId As Integer ) As Integer 
Declare Function AnlActiveSetting Lib "r3cu_AUX_200.DLL" ( ByVal RelayMode As Integer , ByVal RelayTime As Integer ) As Integer 
Declare Function AnlIdle Lib "r3cu_AUX_200.DLL" ( ) As Integer 
'if !defined(__AUX_CONSTS_H)
Public Const uuAUX_CONSTS_H As Integer = 0
'include "r3cUnitErrors.h"
Public Const MAPNAME_SYSTEM_SYNCROBUS As String =        "RP3_SystemSyncrobus"
Public Const MAPNAME_SERVICE_UNIT As String =            "RP3_ServiceUnit"
Public Const OS_INTERNAL As Integer =     1
Public Const OS_INTERFACE As Integer =    2
Public Const LEVEL As Integer =       1
Public Const EDGE As Integer =        2
Public Const MODIN1 As Integer =      &H40
Public Const MODIN2 As Integer =      &H80
Public Const SWITCHING As Integer =   1
Public Const LATCHING As Integer =    2
Public Const T1mS As Integer =        1
Public Const T5mS As Integer =        5
Public Const T10mS As Integer =       10
'endif
'if !defined(__ZBMUMTX_H)
Public Const uuZBMUMTX_H As Integer = 0
'include "r3cuZbmumtxConsts.h"
'ifdef __cplusplus
Declare Function fFTDISerialNumberToCheckSet Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pUnitSn As String ) As Integer 
Declare Function fFTDIInit Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pBaudRate As Integer ) As Integer 
Declare Function fFTDIInitUsingPorts Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pBaudRate As Integer , ByVal pPortMask As Integer ) As Integer 
Declare Function fFTDIHandleGet Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByRef pHandle As Integer ) As Integer 
Declare Function fFTDIClose Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer ) As Integer 
Declare Function fFTDIWrite Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pSizeToTransfer As Integer , ByVal pBuffer As String ) As Integer 
Declare Function fFTDISetIO Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pIOMask As Byte , ByVal pBitMode As Byte ) As Integer 
Declare Function fFTDIClearBuffer Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer ) As Integer 
Declare Function fFTDIRead Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pBuffer As String , ByVal pByteNum As Integer ) As Integer 
Declare Function fFTDIResetDevice Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer ) As Integer 
Declare Function fFTDIResetDeviceUsingPorts Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pPortMask As Integer ) As Integer 
Declare Function fSetMux Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pFtdiPinNum As Integer , ByVal pPEnum As Integer , ByVal pPERw As Integer , ByVal pI2C As Integer ) As Integer 
Declare Function fI2C_InitChannel Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pClockRate As Integer , ByVal pLatencyTimer As Integer , ByVal pOptions As Integer ) As Integer 
Declare Function fI2C_DeviceRead Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pSlaveAddress As Integer , ByVal pSizeToTransfer As Integer , ByVal pBuffer As String , ByRef pSizeTransfered As Integer , ByVal pOptions As Integer ) As Integer 
Declare Function fI2C_DeviceWrite Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pSlaveAddress As Integer , ByVal pSizeToTransfer As Integer , ByVal pBuffer As String , ByRef pSizeTransfered As Integer , ByVal pOptions As Integer ) As Integer 
Declare Function fSPI_InitChannel Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pClockRate As Integer , ByVal pLatencyTimer As Integer , ByVal pOptions As Integer , ByVal pPin As Integer ) As Integer 
Declare Function fSPI_Read Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pBuffer As String , ByVal pSizeToTransfer As Integer , ByRef pSizeTransfered As Integer , ByVal pOptions As Integer ) As Integer 
Declare Function fSPI_Write Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pBuffer As String , ByVal pSizeToTransfer As Integer , ByRef pSizeTransfered As Integer , ByVal pOptions As Integer ) As Integer 
Declare Function fSPI_ReadWrite Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pInBuffer As String , ByVal pOutBuffer As String , ByVal pSizeToTransfer As Integer , ByRef pSizeTransfered As Integer , ByVal pOptions As Integer ) As Integer 
Declare Function fSPI_IsBusy Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByRef pState As Integer ) As Integer 
Declare Function fSPI_ChangeCS Lib "r3cu_ZBMUMTX_200.DLL" ( ByVal pBoardId As Integer , ByVal pConfigOptions As Integer ) As Integer 
'if !defined(__ZBMUMTX_CONSTS_H)
Public Const uuZBMUMTX_CONSTS_H As Integer = 0
Public Const kFT_PORT_A As Integer =     &H01
Public Const kFT_PORT_B As Integer =     &H02
Public Const kFT_PORT_ALL As Integer =   &H03
Public Const kFT_BITBANG_RESET As Integer =     0
Public Const kFT_BITBANG_ASYNC As Integer =     1
Public Const kFT_BITBANG_MPSSE As Integer =     2
Public Const kFT_BITBANG_SYNC As Integer =      4
Public Const kFT_BITBANG_MCU As Integer =       8
Public Const kFT_BITBANG_FAST_OPTO As Integer =16
Public Const kFT_BITBANG_CBUS As Integer =     32        
Public Const kI2C_CLOCK_STANDARD_MODE As Integer =                 0
Public Const kI2C_CLOCK_FAST_MODE As Integer =                     1
Public Const kI2C_CLOCK_FAST_MODE_PLUS As Integer =                2
Public Const kI2C_CLOCK_HIGH_SPEED_MODE As Integer =               3
Public Const kI2C_TRANSFER_OPTIONS_START_BIT		 As Integer =       &H00000001
Public Const kI2C_TRANSFER_OPTIONS_STOP_BIT		 As Integer =       &H00000002
Public Const kI2C_TRANSFER_OPTIONS_BREAK_ON_NACK As Integer =            &H00000004
Public Const kI2C_TRANSFER_OPTIONS_NACK_LAST_BYTE As Integer =           &H00000008
Public Const kI2C_TRANSFER_OPTIONS_FAST_TRANSFER_BYTES As Integer =      &H00000010
Public Const kI2C_TRANSFER_OPTIONS_FAST_TRANSFER_BITS As Integer =       &H00000020
Public Const kI2C_TRANSFER_OPTIONS_FAST_TRANSFER As Integer =            &H00000030
Public Const kI2C_TRANSFER_OPTIONS_NO_ADDRESS As Integer =               &H00000040
Public Const kSPI_CLOCK_STANDARD_MODE As Integer =                 0
Public Const kSPI_CLOCK_FAST_MODE As Integer =                     1
Public Const kSPI_CLOCK_FAST_MODE_PLUS As Integer =                2
Public Const kSPI_CLOCK_HIGH_SPEED_MODE As Integer =               3
Public Const kSPI_TRANSFER_OPTIONS_SIZE_IN_BYTES As Integer =            &H00000000
Public Const kSPI_TRANSFER_OPTIONS_SIZE_IN_BITS As Integer =             &H00000001
Public Const kSPI_TRANSFER_OPTIONS_CHIPSELECT_ENABLE As Integer =        &H00000002
Public Const kSPI_TRANSFER_OPTIONS_CHIPSELECT_DISABLE As Integer =       &H00000004
Public Const kSPI_CONFIG_OPTION_MODE_MASK As Integer =                   &H00000003
Public Const kSPI_CONFIG_OPTION_MODE0 As Integer =                       &H00000000
Public Const kSPI_CONFIG_OPTION_MODE1 As Integer =                       &H00000001
Public Const kSPI_CONFIG_OPTION_MODE2 As Integer =                       &H00000002
Public Const kSPI_CONFIG_OPTION_MODE3 As Integer =                       &H00000003
Public Const kSPI_CONFIG_OPTION_CS_MASK As Integer =                     &H0000001C
Public Const kSPI_CONFIG_OPTION_CS_DBUS3 As Integer =                    &H00000000
Public Const kSPI_CONFIG_OPTION_CS_DBUS4 As Integer =                    &H00000004
Public Const kSPI_CONFIG_OPTION_CS_DBUS5 As Integer =                    &H00000008
Public Const kSPI_CONFIG_OPTION_CS_DBUS6 As Integer =                    &H0000000C
Public Const kSPI_CONFIG_OPTION_CS_DBUS7 As Integer =                    &H00000010
Public Const kSPI_CONFIG_OPTION_CS_ACTIVELOW As Integer =                &H00000020
'endif'ifndef __VRAD_H
Public Const uuVRAD_H As Integer = 0
Public Const MAX_ARRAY_LEN As Integer =                 200
'include "VRad_Types.h"
'include "VRad_Defs.h"
'ifdef __cplusplus
Declare Sub VRAD_LibraryAttach Lib "vradleo200.DLL" ( )   
Declare Sub VRAD_LibraryDetach Lib "vradleo200.DLL" ( )   
Declare Function fvrCheckDongle Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrInitialize Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrInitOnTpgm Lib "vradleo200.DLL" ( ByVal TpgmPath As String ) As Integer 
Declare Function fvrGetTpgmPath Lib "vradleo200.DLL" ( ByVal TpgmPath As String ) As Integer 
Declare Function fvrGetPrintTaskInfo Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetPrintTaskInfo Lib "vradleo200.DLL" ( ByVal PrintTaskInfo As Integer ) As Integer 
Declare Function fvrGetPrintTestInfo Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetPrintTestInfo Lib "vradleo200.DLL" ( ByVal PrintTestInfo As Integer ) As Integer 
Declare Function fvrGetPrintDeviceInfo Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetPrintDeviceInfo Lib "vradleo200.DLL" ( ByVal PrintDeviceInfo As Integer ) As Integer 
Declare Function fvrGetDiscDieOnFailTask Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetDiscDieOnFailTask Lib "vradleo200.DLL" ( ByVal DiscDieOnFailTask As Integer ) As Integer 
Declare Function fvrGetDiscDieOnFailTest Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetDiscDieOnFailTest Lib "vradleo200.DLL" ( ByVal DiscDieOnFailTest As Integer ) As Integer 
Declare Function fvrGetDatalogFlagAlsoForCdColl Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetDatalogFlagAlsoForCdColl Lib "vradleo200.DLL" ( ByVal DatalogFlagAlsoForCdColl As Integer ) As Integer 
Declare Function fvrGetEnableStart Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetEnableStart Lib "vradleo200.DLL" ( ByVal EnableStart As Integer ) As Integer 
Declare Function fvrGetTpgmRelease Lib "vradleo200.DLL" ( ByVal TpgmRelease As String ) As Integer 
Declare Function fvrSetTpgmRelease Lib "vradleo200.DLL" ( ByVal TpgmRelease As String ) As Integer 
Declare Function fvrGetTpgmRemark Lib "vradleo200.DLL" ( ByVal TpgmRemark As String ) As Integer 
Declare Function fvrSetTpgmRemark Lib "vradleo200.DLL" ( ByVal TpgmRemark As String ) As Integer 
Declare Function fvrGetDiesNumber Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetDiesNumber Lib "vradleo200.DLL" ( ByVal DiesNumber As Integer ) As Integer 
Declare Function fvrResetData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetDatalogs Lib "vradleo200.DLL" ( ByVal DatalogNameList As String , ByVal DatalogRelList As String ) As Integer 
Declare Function fvrGetTpgmDatalogInfo Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Release As String , ByVal Path As String , ByVal DefinedTasks As String , ByVal AllTests As String , ByVal FailedTests As String , ByVal TestList As String , ByVal SampleMode As String , ByVal SampleRate As String , ByVal SampleLimit As String , ByVal SamplesNum As String , ByVal SampleType As String , ByVal EnableSummary As String , ByVal FirstDeviceNum As String , ByVal UseTpgmSettings As String ) As Integer 
Declare Function fvrSetTpgmDatalogInfo Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Release As String , ByVal Path As String , ByVal DefinedTasks As String , ByVal AllTests As String , ByVal FailedTests As String , ByVal TestList As String , ByVal SampleMode As String , ByVal SampleRate As String , ByVal SampleLimit As String , ByVal SamplesNum As String , ByVal SampleType As String , ByVal EnableSummary As String , ByVal FirstDeviceNum As String , ByVal UseTpgmSettings As String ) As Integer 
Declare Function fvrSetVRadRelease Lib "vradleo200.DLL" ( ByVal Release As String ) As Integer 
Declare Function fvrSetTpgmStatus Lib "vradleo200.DLL" ( ByVal Status As String ) As Integer 
Declare Function fvrIsLoadDataOnRun Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetTpgmSettings Lib "vradleo200.DLL" ( ByVal CheckForUpdate As String ) As Integer 
Declare Function fvrSetTpgmSettings Lib "vradleo200.DLL" ( ByVal CheckForUpdate As String ) As Integer 
Declare Function fvrSetPatternAdditional Lib "vradleo200.DLL" ( ByVal PatternName As String ) As Integer 
Declare Function fvrGetAllPatternAdditional Lib "vradleo200.DLL" ( ByVal PatternList As String ) As Integer 
Declare Function fvrGetAllPatternUsed Lib "vradleo200.DLL" ( ByVal PatternList As String ) As Integer 
Declare Function fvrGetSystemReset Lib "vradleo200.DLL" ( ByVal SystemReset As String ) As Integer 
Declare Function fvrSetSystemReset Lib "vradleo200.DLL" ( ByVal SystemReset As String ) As Integer 
Declare Function fvrGetConvFactList Lib "vradleo200.DLL" ( ByVal ConvFactList As String ) As Integer 
Declare Function fvrGetPermissions Lib "vradleo200.DLL" ( ByVal DeviceCode As String , ByVal UserLevel As String , ByVal Permissions As String ) As Integer 
Declare Function fvrSetPermissions Lib "vradleo200.DLL" ( ByVal DeviceCode As String , ByVal UserLevel As String , ByVal Permissions As String ) As Integer 
Declare Function fvrGetRevision Lib "vradleo200.DLL" ( ByVal RevNo As String , ByVal Release As String , ByVal vbData As String , ByVal Author As String , ByVal Level1 As String , ByVal Level2 As String , ByVal Level3 As String , ByVal Notes As String ) As Integer 
Declare Function fvrSetRevision Lib "vradleo200.DLL" ( ByVal RevNo As String , ByVal Release As String , ByVal vbData As String , ByVal Author As String , ByVal Level1 As String , ByVal Level2 As String , ByVal Level3 As String , ByVal Notes As String ) As Integer 
Declare Function fvrGetRevisionsNumber Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetRevisionsNumber Lib "vradleo200.DLL" ( ByVal RevisionsNum As String ) As Integer 
Declare Function fvrInitDevicesLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndDevicesLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrInitDevicesSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndDevicesSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetDevice Lib "vradleo200.DLL" ( ByVal Code As String , ByVal Remark As String , ByVal DiesNum As String ) As Integer 
Declare Function fvrSetDevice Lib "vradleo200.DLL" ( ByVal Code As String , ByVal Remark As String , ByVal DiesNum As String ) As Integer 
Declare Function fvrGetInstr Lib "vradleo200.DLL" ( ByVal InstrNo As String , ByVal Name As String , ByVal Remark As String , ByVal EnTplanMask As String , ByVal DieNo As String , ByVal Instruments As String ) As Integer 
Declare Function fvrSetInstr Lib "vradleo200.DLL" ( ByVal InstrNo As String , ByVal Name As String , ByVal Remark As String , ByVal EnTplanMask As String , ByVal DieNo As String , ByVal Instruments As String ) As Integer 
Declare Function fvrGetInstrsNumber Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetInstrsNumber Lib "vradleo200.DLL" ( ByVal InstrsNum As String ) As Integer 
Declare Function fvrGetAllInstrs Lib "vradleo200.DLL" ( ByVal NameList As String , ByVal MaxListLen As Integer ) As Integer 
Declare Function fvrInitBinsLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndBinsLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrInitBinsSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndBinsSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetBin Lib "vradleo200.DLL" ( ByVal Name As String , ByVal SwBin As String , ByVal HwBin As String , ByVal Result As String , ByVal AlertLimit As String ) As Integer 
Declare Function fvrGetBin2 Lib "vradleo200.DLL" ( ByVal Name As String , ByVal SwBin As String , ByVal HwBin As String , ByVal Result As String , ByVal AlertLimit As String , ByVal LBL As String , ByVal HBL As String ) As Integer 
Declare Function fvrGetBin3 Lib "vradleo200.DLL" ( ByVal Name As String , ByVal SwBin As String , ByVal HwBin As String , ByVal Result As String , ByVal AlertLimit As String , ByVal LBL As String , ByVal HBL As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrSetBin Lib "vradleo200.DLL" ( ByVal Name As String , ByVal SwBin As String , ByVal HwBin As String , ByVal Result As String , ByVal AlertLimit As String ) As Integer 
Declare Function fvrSetBin2 Lib "vradleo200.DLL" ( ByVal Name As String , ByVal SwBin As String , ByVal HwBin As String , ByVal Result As String , ByVal AlertLimit As String , ByVal LBL As String , ByVal HBL As String ) As Integer 
Declare Function fvrSetBin3 Lib "vradleo200.DLL" ( ByVal Name As String , ByVal SwBin As String , ByVal HwBin As String , ByVal Result As String , ByVal AlertLimit As String , ByVal LBL As String , ByVal HBL As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrInitPinsLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndPinsLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrInitPinsSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndPinsSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetPin Lib "vradleo200.DLL" ( ByVal Number As String , ByVal Name As String , ByVal vbType As String , ByVal Channels As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrSetPin Lib "vradleo200.DLL" ( ByVal Number As String , ByVal Name As String , ByVal vbType As String , ByVal Channels As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrGetAllPins Lib "vradleo200.DLL" ( ByVal NameList As String , ByVal MaxListLen As Integer ) As Integer 
Declare Function fvrInitPinGroupsLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndPinGroupsLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrInitPinGroupsSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndPinGroupsSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetPinGroup Lib "vradleo200.DLL" ( ByVal Name As String , ByRef PinList As Short , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrSetPinGroup Lib "vradleo200.DLL" ( ByVal Name As String , ByRef PinList As Short , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrGetAllGroups Lib "vradleo200.DLL" ( ByVal NameList As String , ByVal MaxListLen As Integer ) As Integer 
Declare Function fvrGetRelay Lib "vradleo200.DLL" ( ByVal RelayNo As String , ByVal Name As String , ByVal EnTplanMask As String , ByVal DieNo As String , ByVal Relays As String ) As Integer 
Declare Function fvrSetRelay Lib "vradleo200.DLL" ( ByVal RelayNo As String , ByVal Name As String , ByVal EnTplanMask As String , ByVal DieNo As String , ByVal Relays As String ) As Integer 
Declare Function fvrGetRelaysNumber Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetRelaysNumber Lib "vradleo200.DLL" ( ByVal RelaysNum As String ) As Integer 
Declare Function fvrInitTplanLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndTplanLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrInitTplanSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrInitTplanAppend Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndTplanSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetTask Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Enabled As String , ByVal Description As String , ByVal FailBin As String , ByVal PassBin As String , ByVal Datalog As String , ByVal RefreshDie As String , ByVal TaskType As String , ByRef TestsNum As Integer , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrGetTask2 Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Enabled As String , ByVal Description As String , ByVal FailBin As String , ByVal PassBin As String , ByVal Datalog As String , ByVal RefreshDie As String , ByVal Location As String , ByVal TaskType As String , ByRef TestsNum As Integer , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrGetTask3 Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Enabled As String , ByVal Description As String , ByVal FailBin As String , ByVal PassBin As String , ByVal Datalog As String , ByVal RefreshDie As String , ByVal Location As String , ByVal TaskType As String , ByRef TestsNum As Integer , ByVal EnTplanMask As String , ByRef UniqueTaskId As Integer ) As Integer 
Declare Function fvrSetTask Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Enabled As String , ByVal Description As String , ByVal FailBin As String , ByVal PassBin As String , ByVal Datalog As String , ByVal RefreshDie As String , ByVal TaskType As String , ByVal TestsNum As Integer , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrSetTask2 Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Enabled As String , ByVal Description As String , ByVal FailBin As String , ByVal PassBin As String , ByVal Datalog As String , ByVal RefreshDie As String , ByVal Location As String , ByVal TaskType As String , ByVal TestsNum As Integer , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrSetTask3 Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Enabled As String , ByVal Description As String , ByVal FailBin As String , ByVal PassBin As String , ByVal Datalog As String , ByVal RefreshDie As String , ByVal Location As String , ByVal TaskType As String , ByVal TestsNum As Integer , ByVal EnTplanMask As String , ByVal UniqueTaskId As String ) As Integer 
Declare Function fvrGetTest Lib "vradleo200.DLL" ( ByVal Name As String , ByRef TestId As Integer , ByVal FailBin As String , ByVal PassBin As String , ByVal ThLow As String , ByVal ThHigh As String , ByVal Unit As String , ByVal Compare As String , ByVal Testpattern As String , ByVal RefreshDie As String , ByVal Nominal As String , ByVal FormulaThLow As String , ByVal FormulaThHigh As String , ByVal Format As String , ByVal Remark As String , ByVal Enabled As String , ByVal EnTplanMask As String , ByVal TestParams As String ) As Integer 
Declare Function fvrGetTest2 Lib "vradleo200.DLL" ( ByVal Name As String , ByRef TestId As Integer , ByVal FailBin As String , ByVal PassBin As String , ByVal ThLow As String , ByVal ThHigh As String , ByVal Unit As String , ByVal Compare As String , ByVal Testpattern As String , ByVal Datalog As String , ByVal DieResult As String , ByVal RefreshDie As String , ByVal Nominal As String , ByVal FormulaThLow As String , ByVal FormulaThHigh As String , ByVal Format As String , ByVal Remark As String , ByVal Enabled As String , ByVal EnTplanMask As String , ByVal TestParams As String ) As Integer 
Declare Function fvrGetTest3 Lib "vradleo200.DLL" ( ByVal Name As String , ByRef TestId As Integer , ByVal FailBin As String , ByVal PassBin As String , ByVal ThLow As String , ByVal ThHigh As String , ByVal Unit As String , ByVal Compare As String , ByVal Testpattern As String , ByVal Datalog As String , ByVal DieResult As String , ByVal RefreshDie As String , ByVal Nominal As String , ByVal FormulaThLow As String , ByVal FormulaThHigh As String , ByVal Format As String , ByVal Remark As String , ByVal Enabled As String , ByVal EnTplanMask As String , ByVal TestParams As String , ByRef UniqueTestId As Integer ) As Integer 
Declare Function fvrSetTest Lib "vradleo200.DLL" ( ByVal Name As String , ByVal TestId As String , ByVal FailBin As String , ByVal PassBin As String , ByVal ThLow As String , ByVal ThHigh As String , ByVal Unit As String , ByVal Compare As String , ByVal Testpattern As String , ByVal RefreshDie As String , ByVal Nominal As String , ByVal FormulaThLow As String , ByVal FormulaThHigh As String , ByVal Format As String , ByVal Remark As String , ByVal Enabled As String , ByVal EnTplanMask As String , ByVal TestParams As String ) As Integer 
Declare Function fvrSetTest2 Lib "vradleo200.DLL" ( ByVal Name As String , ByVal TestId As String , ByVal FailBin As String , ByVal PassBin As String , ByVal ThLow As String , ByVal ThHigh As String , ByVal Unit As String , ByVal Compare As String , ByVal Testpattern As String , ByVal Datalog As String , ByVal DieResult As String , ByVal RefreshDie As String , ByVal Nominal As String , ByVal FormulaThLow As String , ByVal FormulaThHigh As String , ByVal Format As String , ByVal Remark As String , ByVal Enabled As String , ByVal EnTplanMask As String , ByVal TestParams As String ) As Integer 
Declare Function fvrSetTest3 Lib "vradleo200.DLL" ( ByVal Name As String , ByVal TestId As String , ByVal FailBin As String , ByVal PassBin As String , ByVal ThLow As String , ByVal ThHigh As String , ByVal Unit As String , ByVal Compare As String , ByVal Testpattern As String , ByVal Datalog As String , ByVal DieResult As String , ByVal RefreshDie As String , ByVal Nominal As String , ByVal FormulaThLow As String , ByVal FormulaThHigh As String , ByVal Format As String , ByVal Remark As String , ByVal Enabled As String , ByVal EnTplanMask As String , ByVal TestParams As String , ByVal UniqueTestId As String ) As Integer 
Declare Function fvrGetTestStatus Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetTestStatus Lib "vradleo200.DLL" ( ByVal Status As Integer ) As Integer 
Declare Function fvrGetMeasNum Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetMeasNum Lib "vradleo200.DLL" ( ByVal MeasNum As String ) As Integer 
Declare Function fvrExportTask Lib "vradleo200.DLL" ( ByVal TaskRecord As Integer , ByVal TargetFile As String ) As Integer 
Declare Function fvrImportTask Lib "vradleo200.DLL" ( ByVal SourceFile As String ) As Integer 
Declare Function fvrInitTestParamsLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndTestParamsLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrInitTestParamsSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndTestParamsSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetTestParam Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Value As String , ByVal Unit As String , ByVal Remark As String ) As Integer 
Declare Function fvrSetTestParam Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Value As String , ByVal Unit As String , ByVal Remark As String ) As Integer 
Declare Function fvrGetAllParams Lib "vradleo200.DLL" ( ByVal NameList As String , ByVal MaxListLen As Integer ) As Integer 
Declare Function fvrInitParamsLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndParamsLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrInitParamsSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndParamsSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetParam Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Remark As String , ByVal Value As String , ByVal Unit As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrGetParam2 Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Remark As String , ByVal DebugVal As String , ByVal ProdVal As String , ByVal Unit As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrSetParam Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Remark As String , ByVal Value As String , ByVal Unit As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrSetParam2 Lib "vradleo200.DLL" ( ByVal Name As String , ByVal Remark As String , ByVal DebugVal As String , ByVal ProdVal As String , ByVal Unit As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrGetArray Lib "vradleo200.DLL" ( ByVal ArrayNo As String , ByVal Name As String , ByVal Remark As String , ByVal vbType As String , ByVal Values As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrSetArray Lib "vradleo200.DLL" ( ByVal ArrayNo As String , ByVal Name As String , ByVal Remark As String , ByVal vbType As String , ByVal Values As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrGetArraysNumber Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetArraysNumber Lib "vradleo200.DLL" ( ByVal ArraysNum As String ) As Integer 
Declare Function fvrGetProbeCard Lib "vradleo200.DLL" ( ByVal ProbeCardNo As String , ByVal Name As String , ByVal DiesNum As String , ByVal DiesPos As String ) As Integer 
Declare Function fvrSetProbeCard Lib "vradleo200.DLL" ( ByVal ProbeCardNo As String , ByVal Name As String , ByVal DiesNum As String , ByVal DiesPos As String ) As Integer 
Declare Function fvrGetProbeCardsNumber Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetProbeCardsNumber Lib "vradleo200.DLL" ( ByVal ProbeCardsNum As String ) As Integer 
Declare Function fvrGetAllProbeCards Lib "vradleo200.DLL" ( ByVal ProbeCards As String ) As Integer 
Declare Function fvrGetDeviceProbeCard Lib "vradleo200.DLL" ( ByVal DeviceCode As String , ByVal ProbeCardName As String ) As Integer 
Declare Function fvrSetDeviceProbeCard Lib "vradleo200.DLL" ( ByVal DeviceCode As String , ByVal ProbeCardName As String ) As Integer 
Declare Function fvrGetDiesXY Lib "vradleo200.DLL" ( ByVal DeviceCode As Short , ByRef XPos As Integer , ByRef YPos As Integer ) As Integer 
Declare Function fvrGetPowerSupply Lib "vradleo200.DLL" ( ByVal StepNo As String , ByVal PSUListOn As String , ByVal PSUListOff As String , ByVal DelayOn As String , ByVal DelayOff As String ) As Integer 
Declare Function fvrSetPowerSupply Lib "vradleo200.DLL" ( ByVal StepNo As String , ByVal PSUListOn As String , ByVal PSUListOff As String , ByVal DelayOn As String , ByVal DelayOff As String ) As Integer 
Declare Function fvrSetPowerSupplyOnUsed Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetPowerSupplyOffUsed Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetSetup Lib "vradleo200.DLL" ( ByVal SetupType As Integer , ByVal SetupNo As String , ByVal Param As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrSetSetup Lib "vradleo200.DLL" ( ByVal SetupType As Integer , ByVal SetupNo As String , ByVal Param As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrGetSetupsNumber Lib "vradleo200.DLL" ( ByVal SetupType As Integer ) As Integer 
Declare Function fvrSetSetupsNumber Lib "vradleo200.DLL" ( ByVal SetupType As Integer , ByVal SetupsNum As String ) As Integer 
Declare Function fvrGetParamValue Lib "vradleo200.DLL" ( ByVal Param As String ) As Integer 
Declare Function fvrGetUnitList Lib "vradleo200.DLL" ( ByVal Units As String , ByRef UnitList As Integer ) As Integer 
Declare Function fvrInitCharactLoad Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrInitCharactSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEndCharactSave Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetCharactTest Lib "vradleo200.DLL" ( ByVal SetupName As String , ByVal TestId As String , ByVal Title As String , ByVal Mode As String , ByVal Pin As String , ByVal XParam As String , ByVal XRemark As String , ByVal XFrom As String , ByVal XTo As String , ByVal XStep As String , ByVal YParam As String , ByVal YRemark As String , ByVal YFrom As String , ByVal YTo As String , ByVal YStep As String , ByVal ZRemark As String , ByVal ZFrom As String , ByVal ZTo As String , ByVal ZStep As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrSetCharactTest Lib "vradleo200.DLL" ( ByVal SetupName As String , ByVal TestId As String , ByVal Title As String , ByVal Mode As String , ByVal Pin As String , ByVal XParam As String , ByVal XRemark As String , ByVal XFrom As String , ByVal XTo As String , ByVal XStep As String , ByVal YParam As String , ByVal YRemark As String , ByVal YFrom As String , ByVal YTo As String , ByVal YStep As String , ByVal ZRemark As String , ByVal ZFrom As String , ByVal ZTo As String , ByVal ZStep As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrGetDPATTest Lib "vradleo200.DLL" ( ByVal TestNo As String , ByVal TestId As String , ByVal SigmaFactor As String , ByVal Bin As String , ByVal RejectLimit As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrSetDPATTest Lib "vradleo200.DLL" ( ByVal TestNo As String , ByVal TestId As String , ByVal SigmaFactor As String , ByVal Bin As String , ByVal RejectLimit As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrGetDPATTestsNumber Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetDPATTestsNumber Lib "vradleo200.DLL" ( ByVal TestsNum As String ) As Integer 
Declare Function fvrGetQAOTFDevice Lib "vradleo200.DLL" ( ByVal DeviceNo As String , ByVal DeviceName As String , ByVal FailBin As String , ByVal SigmaRate As String , ByVal SampleLimit As String , ByVal SampleType As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrSetQAOTFDevice Lib "vradleo200.DLL" ( ByVal DeviceNo As String , ByVal DeviceName As String , ByVal FailBin As String , ByVal SigmaRate As String , ByVal SampleLimit As String , ByVal SampleType As String , ByVal EnTplanMask As String ) As Integer 
Declare Function fvrGetQAOTFDevicesNumber Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetQAOTFDevicesNumber Lib "vradleo200.DLL" ( ByVal DevicesNum As String ) As Integer 
Declare Function fvrMapGlobalMemory Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrUnmapGlobalMemory Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadFactorList Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadDeviceData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadBinData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadPinData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadPinGroupData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadRelayData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadInstrData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadTaskData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadTestData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadTestParamsData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadParamData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadDieData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrLoadTpgmData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrUnloadTpgmData Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetDeviceList Lib "vradleo200.DLL" ( ByVal DeviceList As String ) As Integer 
Declare Function fvrGetDeviceId Lib "vradleo200.DLL" ( ByVal DeviceName As String ) As Short 
Declare Function fvrGetDeviceNoFromName Lib "vradleo200.DLL" ( ByVal DeviceName As String ) As Short 
Declare Function fvrGetDeviceNoFromId Lib "vradleo200.DLL" ( ByVal DeviceId As Short ) As Short 
Declare Function fvrGetTaskList Lib "vradleo200.DLL" ( ByVal TaskList As String , ByRef EnTaskList As Byte ) As Integer 
Declare Function fvrSelectDevice Lib "vradleo200.DLL" ( ByVal DeviceNo As Integer ) As Integer 
Declare Function fvrLoadSetups Lib "vradleo200.DLL" ( ByVal DeviceNo As Integer ) As Integer 
Declare Function fvrEnableDie Lib "vradleo200.DLL" ( ByVal DieNo As Integer , ByVal Enable As Integer ) As Integer 
Declare Function fvrSetDiePresent Lib "vradleo200.DLL" ( ByVal DieNo As Integer , ByVal Present As Integer ) As Integer 
Declare Function fvrSetTpgmPinCh Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrEnableTask Lib "vradleo200.DLL" ( ByVal TaskName As String , ByVal Enable As Integer ) As Integer 
Declare Function fvrGetTestList Lib "vradleo200.DLL" ( ByRef TestList As Integer , ByRef TestsNum As Integer ) As Integer 
Declare Function fvrGetTestPosition Lib "vradleo200.DLL" ( ByVal TestNo As Integer , ByRef TaskPos As Integer , ByRef TestPos As Integer ) As Integer 
Declare Function fvrGetTestPositionFromTask Lib "vradleo200.DLL" ( ByVal TaskPos As Integer , ByVal TestNo As Integer , ByRef TestPos As Integer ) As Integer 
Declare Function fvrGetThr Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByRef ThrL As Double , ByRef ThrH As Double , ByRef ThrLInfo As Byte , ByRef ThrHInfo As Byte ) As Integer 
Declare Function fvrGetNominalValue Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByRef NominalValue As Double ) As Integer 
Declare Function fvrSetNominalValue Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByVal NominalValue As Double ) As Integer 
Declare Function fvrGetMeasUnit Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByVal MeasUnit As String ) As Integer 
Declare Function fvrSetMeasUnit Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByVal MeasUnit As String ) As Integer 
Declare Function fvrGetCompMode Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByRef CompMode As Byte ) As Integer 
Declare Function fvrGetFormat Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByVal Format As String ) As Integer 
Declare Function fvrGetTaskName Lib "vradleo200.DLL" ( ByVal TaskPos As Integer , ByVal TaskName As String ) As Integer 
Declare Function fvrGetTaskType Lib "vradleo200.DLL" ( ByVal TaskPos As Integer , ByRef TaskType As Byte ) As Integer 
Declare Function fvrGetTaskDatalog Lib "vradleo200.DLL" ( ByVal TaskPos As Integer , ByRef TaskDatalog As Byte ) As Integer 
Declare Function fvrGetTestName Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByVal TestName As String ) As Integer 
Declare Function fvrGetTestBins Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByRef PassBin As Short , ByRef PassBinType As Byte , ByRef FailBin As Short , ByRef FailBinType As Byte ) As Integer 
Declare Function fvrGetBinType Lib "vradleo200.DLL" ( ByVal SwBin As Short , ByRef BinType As Byte ) As Integer 
Declare Function fvrGetBinData Lib "vradleo200.DLL" ( ByVal SwBin As Short , ByRef BinType As Byte , ByVal BinName As String , ByRef Alert As Integer ) As Integer 
Declare Function fvrGetTaskRefreshDie Lib "vradleo200.DLL" ( ByVal TaskPos As Integer , ByRef RefreshDie As Integer ) As Integer 
Declare Function fvrGetTestRefreshDie Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByRef RefreshDie As Integer ) As Integer 
Declare Function fvrGetHwBin Lib "vradleo200.DLL" ( ByVal SwBin As Short , ByRef HwBin As Short ) As Integer 
Declare Function fvrGetTasksNum Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetTestsNum Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetActualDiesNum Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetMaxDiesNum Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetRelayList Lib "vradleo200.DLL" ( ByVal RelayName As String , ByRef DieList As Short , ByRef Relays As Short ) As Integer 
Declare Function fvrGetBinName Lib "vradleo200.DLL" ( ByVal SwBin As Short , ByVal BinName As String ) As Integer 
Declare Function fvrGetCharactSetupsList Lib "vradleo200.DLL" ( ByVal SetupsList As String ) As Integer 
'  'ifdef __cplusplus
'  int WINAPI fvrGetCharactSetupTests (char *SetupName, DWORD *TestIndexList=NULL, DWORD *TestIdList=NULL);
Declare Function fvrGetCharactSetupTests Lib "vradleo200.DLL" ( ByVal SetupName As String , ByRef TestIndexList As Integer , ByRef TestIdList As Integer ) As Integer 
'  BOOL WINAPI fvrGetCharactParams (int TestIdx, typCHARACTDATA *Data);
Declare Function fvrIsTestEnabled Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestPos As Integer ) As Integer 
Declare Function fvrFillDieXY Lib "vradleo200.DLL" ( ByVal X1 As Integer , ByVal Y1 As Integer , ByVal XRightToLeft As Integer , ByVal YUpToDown As Integer , ByRef XPos As Integer , ByRef YPos As Integer ) As Integer 
Declare Function fvrGetActivePins Lib "vradleo200.DLL" ( ByRef PinList As Short ) As Integer 
Declare Function fvrGetActivePinGroups Lib "vradleo200.DLL" ( ByVal PinGroupList As String ) As Integer 
Declare Function fvrGetPinList Lib "vradleo200.DLL" ( ByVal PinList As String , ByRef wPinList As Short ) As Integer 
Declare Function fvrGetPinType Lib "vradleo200.DLL" ( ByVal PinNo As Short , ByRef PinType As Byte ) As Integer 
Declare Function fvrGetFactor Lib "vradleo200.DLL" ( ByVal Unit As String ) As Double 
Declare Function fvrGetBinAlarms Lib "vradleo200.DLL" ( ByVal SwBin As Short , ByRef LBL As Byte , ByRef HBL As Byte ) As Integer 
Declare Function fvrGetBinAlertLimit Lib "vradleo200.DLL" ( ByVal SwBin As Short , ByRef AlertLimit As Integer , ByRef IsPercentage As Integer ) As Integer 
Declare Function fvrIsDPATEnabled Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetDPATSampleRate Lib "vradleo200.DLL" ( ) As Short 
'  'ifdef __cplusplus
'  DWORD WINAPI fvrGetDPATTestsInfo (typDPATTEST *DPATTest = NULL);
'  DWORD WINAPI fvrGetDPATTestsInfo (typDPATTEST *DPATTest);
Declare Function fvrSetDoubleLocalParameter Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByVal ParamName As String , ByVal Value As Double ) As Integer 
Declare Function fvrIsQAOTFEnabled Lib "vradleo200.DLL" ( ) As Integer 
'  BOOL WINAPI fvrGetQAOTFDeviceInfo (typQAOTFDEVICE *QAOTFDevice);
Declare Function fvrSetCustomProbeCardLayout Lib "vradleo200.DLL" ( ByVal DiesNum As Integer , ByRef XPos As Integer , ByRef YPos As Integer ) As Integer 
Declare Function fvrGetPatternName Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByVal PattName As String ) As Integer 
Declare Function fvrGetTaskLocation Lib "vradleo200.DLL" ( ByVal TaskPos As Integer , ByRef Location As Short ) As Integer 
Declare Function fvrIsDeviceVisible Lib "vradleo200.DLL" ( ByVal DeviceId As Short ) As Integer 
Declare Function fvrGetProbeCardLayout Lib "vradleo200.DLL" ( ByRef XPos As Integer , ByRef YPos As Integer ) As Integer 
Declare Function fvrGetTestDatalog Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByRef Datalog As Integer ) As Integer 
Declare Function fvrGetTestDieResult Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByRef DieResult As Integer ) As Integer 
Declare Function fvrGetTpgmGenPath Lib "vradleo200.DLL" ( ByVal TpgmGenPath As String ) As Integer 
Declare Function fvrGetLibExpPath Lib "vradleo200.DLL" ( ByVal LibExpPath As String ) As Integer 
Declare Function fvrGetTestRemark Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByVal TestRemark As String ) As Integer 
Declare Function fvrGetTaskDescriptionList Lib "vradleo200.DLL" ( ByVal TaskList As String , ByRef EnTaskList As Byte ) As Integer 
Declare Function GetDieList Lib "vradleo200.DLL" ( ByRef DieList As Short ) As Integer 
Declare Function GetSiteList Lib "vradleo200.DLL" ( ByRef SiteList As Short ) As Integer 
Declare Function IsDieEnabled Lib "vradleo200.DLL" ( ByVal DieNum As Short ) As Integer 
Declare Function IsDiePresent Lib "vradleo200.DLL" ( ByVal DieNum As Short ) As Integer 
Declare Function IsTaskEnabled Lib "vradleo200.DLL" ( ByVal TaskNo As Integer ) As Integer 
Declare Function IsTestEnabled Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer ) As Integer 
Declare Function IsDieListEmpty Lib "vradleo200.DLL" ( ) As Integer 
Declare Function GetPinNumber Lib "vradleo200.DLL" ( ByVal PinName As String ) As Short 
Declare Function GetPinName Lib "vradleo200.DLL" ( ByVal PinNo As Short , ByVal PinName As String ) As Integer 
Declare Function GetPinList Lib "vradleo200.DLL" ( ByVal PinGroupName As String , ByRef PinList As Short ) As Integer 
Declare Function GetInstrumentList Lib "vradleo200.DLL" ( ByVal InstrListName As String , ByRef InstrList As Integer ) As Integer 
Declare Function GetThresholds Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByRef ThrLow As Double , ByRef ThrHigh As Double ) As Integer 
Declare Function SetThresholds Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByVal ThrLow As Double , ByVal ThrHigh As Double ) As Integer 
Declare Function SetHighThr Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByVal ThrHigh As Double ) As Integer 
Declare Function SetLowThr Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByVal ThrLow As Double ) As Integer 
Declare Function GetParameter Lib "vradleo200.DLL" ( ByVal ParamName As String , ByRef ParamValue As Double ) As Integer 
Declare Function GetCharParameter Lib "vradleo200.DLL" ( ByVal ParamName As String , ByVal ParamValue As String ) As Integer 
Declare Function GetCharParameterV2 Lib "vradleo200.DLL" ( ByVal ParamName As String , ByVal ParamValue As String ) As Integer 
Declare Function GetNominalValue Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByRef Nominal As Double ) As Integer 
Declare Function SetNominalValue Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByVal Nominal As Double ) As Integer 
Declare Function GetSelectedDevice Lib "vradleo200.DLL" ( ByRef DeviceSel As Integer , ByVal Remark As String ) As Integer 
Declare Function GetDoubleLocalParameter Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByVal ParamName As String ) As Double 
Declare Function GetBoolLocalParameter Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByVal ParamName As String ) As Integer 
Declare Function GetIntLocalParameter Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByVal ParamName As String ) As Integer 
Declare Function GetCharLocalParameter Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByVal ParamName As String ) As String 
Declare Function GetCharLocalParameterV2 Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByVal ParamName As String , ByVal Parameter As String ) As Integer 
Declare Function GetLocalParameterFactorUnit Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByVal ParamName As String ) As Double 
Declare Function IsIgnoreRefreshDie Lib "vradleo200.DLL" ( ByVal TaskNo As Integer ) As Integer 
'  'ifdef __cplusplus
'  DWORD WINAPI GetTaskTestList (int TaskNo, DWORD *TestList = NULL);
Declare Function GetTaskTestList Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByRef TestList As Integer ) As Integer 
Declare Function GetTTestName Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByVal TestName As String ) As Integer 
'  int WINAPI FillArrayData (char *ArrayName, void *RetArray);
Declare Function FillArrayDataString Lib "vradleo200.DLL" ( ByVal ArrayName As String , ByVal RetArray As String ) As Integer 
Declare Function FillArrayDataInt Lib "vradleo200.DLL" ( ByVal ArrayName As String , ByRef RetArray As Integer ) As Integer 
Declare Function FillArrayDataDouble Lib "vradleo200.DLL" ( ByVal ArrayName As String , ByRef RetArray As Double ) As Integer 
Declare Function FillArrayDataByte Lib "vradleo200.DLL" ( ByVal ArrayName As String , ByRef RetArray As Byte ) As Integer 
Declare Function FillArrayDataWord Lib "vradleo200.DLL" ( ByVal ArrayName As String , ByRef RetArray As Short ) As Integer 
Declare Function LoadUsedTestpatterns Lib "vradleo200.DLL" ( ) As Integer 
Declare Function IsTestIgnoreRefreshDie Lib "vradleo200.DLL" ( ByVal TestNo As Integer ) As Integer 
Declare Function GetThresholdsConvFact Lib "vradleo200.DLL" ( ByVal TaskNo As Integer , ByVal TestId As Integer , ByRef ConvFact As Double ) As Integer 
Declare Function TpgmGetDeviceList Lib "vradleo200.DLL" ( ByVal DeviceList As String ) As Integer 
Declare Function TpgmGetDeviceId Lib "vradleo200.DLL" ( ByVal DeviceName As String ) As Short 
Declare Function TpgmGetTaskList Lib "vradleo200.DLL" ( ByVal TaskList As String , ByRef EnTaskList As Byte ) As Integer 
Declare Function TpgmEnableTask Lib "vradleo200.DLL" ( ByVal TaskName As String , ByVal Enable As Integer ) As Integer 
Declare Function TpgmGetTestList Lib "vradleo200.DLL" ( ByRef TestList As Integer , ByRef TestsNum As Integer ) As Integer 
Declare Function TpgmGetTestHandle Lib "vradleo200.DLL" ( ByVal TestNo As Integer , ByRef TaskHnd As Integer , ByRef TestHnd As Integer ) As Integer 
Declare Function TpgmGetThr Lib "vradleo200.DLL" ( ByVal TestHnd As Integer , ByRef ThrL As Double , ByRef ThrH As Double , ByRef ThrLInfo As Byte , ByRef ThrHInfo As Byte ) As Integer 
Declare Function TpgmGetNominalValue Lib "vradleo200.DLL" ( ByVal TestHnd As Integer , ByRef NominalValue As Double ) As Integer 
Declare Function TpgmSetNominalValue Lib "vradleo200.DLL" ( ByVal TestHnd As Integer , ByVal NominalValue As Double ) As Integer 
Declare Function TpgmGetMeasUnit Lib "vradleo200.DLL" ( ByVal TestHnd As Integer , ByVal MeasUnit As String ) As Integer 
Declare Function TpgmSetMeasUnit Lib "vradleo200.DLL" ( ByVal TestHnd As Integer , ByVal MeasUnit As String ) As Integer 
Declare Function TpgmGetCompMode Lib "vradleo200.DLL" ( ByVal TestHnd As Integer , ByRef CompMode As Byte ) As Integer 
Declare Function TpgmGetFormat Lib "vradleo200.DLL" ( ByVal TestHnd As Integer , ByVal Format As String ) As Integer 
Declare Function TpgmGetTaskName Lib "vradleo200.DLL" ( ByVal TaskHnd As Integer , ByVal TaskName As String ) As Integer 
Declare Function TpgmGetTaskType Lib "vradleo200.DLL" ( ByVal TaskHnd As Integer , ByRef TaskType As Byte ) As Integer 
Declare Function TpgmGetTaskDatalog Lib "vradleo200.DLL" ( ByVal TaskHnd As Integer , ByRef TaskDatalog As Byte ) As Integer 
Declare Function TpgmGetTestName Lib "vradleo200.DLL" ( ByVal TestHnd As Integer , ByVal TestName As String ) As Integer 
Declare Function TpgmGetTestBins Lib "vradleo200.DLL" ( ByVal TestHnd As Integer , ByRef PassBin As Short , ByRef PassBinType As Byte , ByRef FailBin As Short , ByRef FailBinType As Byte ) As Integer 
Declare Function TpgmGetBinType Lib "vradleo200.DLL" ( ByVal SwBin As Short , ByRef BinType As Byte ) As Integer 
Declare Function TpgmGetTaskRefreshDie Lib "vradleo200.DLL" ( ByVal TaskHnd As Integer , ByRef RefreshDie As Integer ) As Integer 
Declare Function TpgmGetHwBin Lib "vradleo200.DLL" ( ByVal SwBin As Short , ByRef HwBin As Short ) As Integer 
Declare Function TpgmGetBinData Lib "vradleo200.DLL" ( ByVal SwBin As Short , ByRef BinType As Byte , ByVal BinName As String , ByRef Alert As Integer ) As Integer 
Declare Function TpgmGetTasksNum Lib "vradleo200.DLL" ( ) As Integer 
Declare Function TpgmGetTestsNum Lib "vradleo200.DLL" ( ) As Integer 
Declare Function TpgmGetActualDiesNum Lib "vradleo200.DLL" ( ) As Integer 
Declare Function TpgmGetMaxDiesNum Lib "vradleo200.DLL" ( ) As Integer 
Declare Function TpgmSetDiePresent Lib "vradleo200.DLL" ( ByVal DieNo As Integer , ByVal Present As Integer ) As Integer 
Declare Function TpgmGetRelayList Lib "vradleo200.DLL" ( ByVal RelayName As String , ByRef DieList As Short , ByRef Relays As Short ) As Integer 
Declare Function TpgmGetBinName Lib "vradleo200.DLL" ( ByVal SwBin As Short , ByVal BinName As String ) As Integer 
Declare Function TpgmGetActivePins Lib "vradleo200.DLL" ( ByRef PinList As Short ) As Integer 
Declare Function TpgmGetActivePinGroups Lib "vradleo200.DLL" ( ByVal PinGroupList As String ) As Integer 
Declare Function TpgmGetPinList Lib "vradleo200.DLL" ( ByVal Pins As String , ByRef wPinList As Short ) As Integer 
Declare Function TpgmGetPinType Lib "vradleo200.DLL" ( ByVal PinNo As Short , ByRef PinType As Byte ) As Integer 
Declare Function TpgmGetFactor Lib "vradleo200.DLL" ( ByVal Unit As String ) As Double 
Declare Function TpgmGetBinAlarms Lib "vradleo200.DLL" ( ByVal SwBin As Short , ByRef LBL As Byte , ByRef HBL As Byte ) As Integer 
Declare Function TpgmGetPatternName Lib "vradleo200.DLL" ( ByVal TestHnd As Integer , ByVal PattName As String ) As Integer 
Declare Function TpgmGetTestDatalog Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByRef Datalog As Integer ) As Integer 
Declare Function TpgmGetTestDieResult Lib "vradleo200.DLL" ( ByVal TestPos As Integer , ByRef DieResult As Integer ) As Integer 
Declare Function TpgmEnableSite Lib "vradleo200.DLL" ( ByVal SiteNo As Integer , ByVal Enable As Integer ) As Integer 
Declare Function TpgmSetSitePresent Lib "vradleo200.DLL" ( ByVal SiteNo As Integer , ByVal Present As Integer ) As Integer 
Declare Function fvrRunLibExp Lib "vradleo200.DLL" ( ByVal UniqueTaskId As Integer , ByVal UniqueTestId As Integer ) As Integer 
Declare Function fvrGetTestFunction Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal FuncId As Integer , ByVal FuncName As String , ByVal FuncParameters As String , ByVal FuncFilePath As String ) As Integer 
Declare Function fvrSetTestFunction Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal FuncId As Integer , ByVal FuncName As String , ByVal FuncParameters As String , ByVal FuncFilePath As String ) As Integer 
Declare Function fvrGetAllTestFunctions Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal FuncName As String ) As Integer 
Declare Function fvrGetTempTestFunction Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal FuncId As Integer , ByVal FuncName As String , ByVal FuncParameters As String , ByVal FuncFilePath As String ) As Integer 
Declare Function fvrSetTempTestFunction Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal FuncId As Integer , ByVal FuncName As String , ByVal FuncParameters As String , ByVal FuncFilePath As String ) As Integer 
Declare Function fvrGetAllTempTestFunctions Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal FuncName As String ) As Integer 
Declare Function fvrClearTestFunctions Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer ) As Integer 
Declare Function fvrClearTempTestFunctions Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer ) As Integer 
Declare Function fvrGetTestFunctionName Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal FuncId As Integer , ByVal FuncName As String ) As Integer 
Declare Function fvrGetTempTestFunctionName Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal FuncId As Integer , ByVal FuncName As String ) As Integer 
Declare Function fvrGetTestFunctionParameters Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal FuncId As Integer , ByVal FuncParameters As String ) As Integer 
Declare Function fvrGetTempTestFunctionParameters Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal FuncId As Integer , ByVal FuncParameters As String ) As Integer 
Declare Function fvrSetTestFunctionParameters Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal FuncId As Integer , ByVal FuncParameters As String ) As Integer 
Declare Function fvrSetTempTestFunctionParameters Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal FuncId As Integer , ByVal FuncParameters As String ) As Integer 
Declare Function fvrGetUniqueTaskId Lib "vradleo200.DLL" ( ByRef UniqueTaskId As Integer ) As Integer 
Declare Function fvrGetUniqueTestId Lib "vradleo200.DLL" ( ByRef UniqueTestId As Integer ) As Integer 
Declare Function fvrGetGenTaskName Lib "vradleo200.DLL" ( ByVal UniqueTaskId As Integer , ByVal GenTaskName As String ) As Integer 
Declare Function fvrGetGenTestId Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByRef GenTestId As Integer ) As Integer 
Declare Function fvrSetGenTaskName Lib "vradleo200.DLL" ( ByVal UniqueTaskId As Integer , ByVal GenTaskName As String ) As Integer 
Declare Function fvrSetGenTestId Lib "vradleo200.DLL" ( ByVal UniqueTestId As Integer , ByVal GenTestId As Integer ) As Integer 
Declare Function fvrGetAllGenTestId Lib "vradleo200.DLL" ( ByVal pBuffer As String , ByVal pBufSize As Integer ) As Integer 
Declare Function fvrClearAllGenTaskName Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrClearAllGenTestId Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrGetFuncPrintMode Lib "vradleo200.DLL" ( ) As Integer 
Declare Function fvrSetFuncPrintMode Lib "vradleo200.DLL" ( ByVal NewMode As Integer ) As Integer 
Declare Function fvrCreateProcess Lib "vradleo200.DLL" ( ByVal ProcessToRun As String ) As Integer 
Declare Function fvrGetTaskNameV2 Lib "vradleo200.DLL" ( ByVal TaskPos As Integer , ByVal TaskName As String ) As Integer 
Declare Function TpgmGetTaskNameV2 Lib "vradleo200.DLL" ( ByVal TaskHnd As Integer , ByVal TaskName As String ) As Integer 
Declare Function fvrBreakOnTaskSet Lib "vradleo200.DLL" ( ByVal TaskPos As Integer , ByVal FlagSet As Integer ) As Integer 
Declare Function fvrIsBreakOnTaskSet Lib "vradleo200.DLL" ( ByVal TaskPos As Integer ) As Integer 
Declare Function fvrBreakOnTestSet Lib "vradleo200.DLL" ( ByVal TaskPos As Integer , ByVal FlagSet As Integer ) As Integer 
Declare Function fvrIsBreakOnTestSet Lib "vradleo200.DLL" ( ByVal TestPos As Integer ) As Integer 
'ifndef r3csFunctionalVradXlsH
Public Const r3csFunctionalVradXlsH As Integer = 0
Public Const MAX_BINS As Integer = 64
Public Const MAX_DIES As Integer = 256
Public Const MAX_PINS As Integer = 512
'ifdef __cplusplus
Declare Function GetBreakTaskId Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function GetBreakTestId Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function IsTaskProfile Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function IsTestProfile Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function IsPrintTaskInfo Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function IsDiscDiesAtEndTask Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function IsDiscDiesAtEndTest Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function IsDoNotDisconnectDlogDies Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function GetSoftwareSyncType Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Sub ExecuteSoftwareSync Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal dwTestNumber As Integer )   
Declare Function ShmooIsTestEnabled Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal dwTestNumber As Integer ) As Integer 
Declare Function ShmooCmdBeginTest Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function ShmooIncreaseParameters Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Sub ShmooClearTestMeas Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal dwTestNumber As Integer )   
Declare Sub ShmooCmdSendMeasure Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Sub ShmooCmdEndTest Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Sub ShmooCmdTpgmReady Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Function GetFirstDeviceNum Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function IsDieToDatalog Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal wDieNum As Short ) As Integer 
Declare Sub GetTaskExecutionInfo Lib "R3CS_FUNCTIONAL_200.DLL" ( ByRef TaskId As Integer , ByVal TaskName As String , ByRef TestNum As Integer , ByVal TestName As String )   
Declare Sub SetTaskExecutionInfo Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TaskId As Integer , ByVal TaskName As String , ByVal TestNum As Integer , ByVal TestName As String )   
Declare Sub SetTpgmResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal wTpgmResult As Short )   
Declare Sub SetLedResultStatus Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal wTpgmResult As Short )   
Declare Function GetHndProbType Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Short 
Declare Function GetTimeElapsed Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Double 
Declare Function IsPrintTestInfo Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function IsTestToPrint Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal TestResult As Integer ) As Integer 
Declare Function GetPinDieChannels Lib "R3CS_FUNCTIONAL_200.DLL" ( ByRef PinList As Short , ByRef DieList As Short , ByRef ChList As Short ) As Integer 
Declare Function GetPinDieChannel Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal PinNr As Short , ByVal DieNr As Short ) As Short 
'  char *WINAPI Res2S (int Result, char *StrResult);
Declare Sub StorePrintFuncValue Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNumber As Integer , ByVal TestNumberInTask As Integer , ByVal UniqueTestId As Integer , ByVal TaskNumber As Integer , ByVal DrawingReference As String , ByVal Remark As String , ByVal TestResult As Integer , ByVal MeasuredValue As Double , ByVal ThrLow As Double , ByVal ThrHigh As Double , ByVal MeasFactor As String , ByVal MeasUnit As String , ByVal MeasuredValueStr As String , ByVal ThrLowStr As String , ByVal ThrHighStr As String , ByVal TpListStr As String , ByVal AltMsg As String , ByVal SiteNum As Integer , ByVal pStoreToReport As Integer , ByVal pStoreToDatalog As Integer , ByVal pStoreToCdcoll As Integer )   
''ifdef __cplusplus
'}
'endif

Public Const SWSYNC_NONE As Integer =     0
Public Const SWSYNC_TEST_NUM As Integer = 1

'----------------------------------------------------------------------------
' Compare results
'
Public Const CPASS As Integer =                     0
Public Const CLFAIL As Integer =                    1
Public Const CHFAIL As Integer =                    2
Public Const CFAIL As Integer =                     3
Public Const CERROR As Integer =                    4

'----------------------------------------------------------------------------
' Compare Mode
'
Public Const INLIMITS As Integer =                  1
Public Const OUTLIMITS As Integer =                 2
Public Const INLIMITS_FAIL As Integer =             3
Public Const OUTLIMITS_FAIL As Integer =            4
Public Const CMP_NOERROR As Integer =               10


Public Const TPGM_PASS As Integer = 0
Public Const TPGM_FAIL As Integer = 1
Public Const TPGM_STOP As Integer = 2

Public Const PROC_CONTROL_PANEL As Integer =0

Public Const HNDPRB_NONE As Integer =   0
Public Const HNDPRB_PROBER As Integer = 1
Public Const HNDPRB_STRIP As Integer =  2





'Public Const MAX_PATTERN_NAME_LEN As Integer = 128
'Public Const MAX_CHANNELS As Integer = 256

''typedef 'struct _PATTERN_ERR
''{
'  DWORD InstrId;                            // In'struction identifier
'  char PattName [MAX_PATTERN_NAME_LEN];     // Pattern Name
'  WORD PattNumber;                          // Pattern Number (internal id.)
'  int Status;                               // Execution status (0=Ok)
'  WORD NumErrors;                           // Number errors
'  DWORD StepNumber [MAX_CHANNELS];          // Array of failed steps
'  WORD PinNumber [MAX_CHANNELS];            // Array of failed pins
'  WORD DieNumber [MAX_CHANNELS];            // Failed dies (for each pin)
'  WORD ChnNumber [MAX_CHANNELS];            // Failed chns.(for each pin)
'  BYTE ErrorType [MAX_CHANNELS];            // Pin Error Type: PIN_LEVEL_ERR, PIN_FUNCT_ERR
''} PATTERN_ERR, *PATTERN_ERR_PTR;
'endif
'if !defined(__TPGMCONTROL_H)
Public Const uuTPGMCONTROL_H As Integer = 0
'include "TpgmCtrl_Defs.h"
'include "TpgmCtrl_Types.h"
'ifdef __cplusplus
Declare Function SetControlTplanFunctions Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal tplan As Integer , ByVal tplanInitialize As Integer , ByVal tplanEnd As Integer ) As Integer 
Declare Function SetOptButtFunction Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal OptButt As Integer , ByVal ButtonId As Integer ) As Integer 
Declare Function SetDiscFailDieFunction Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal DiscFailedDies As Integer , ByRef FailDieList As Short ) As Integer 
Declare Function SetUpdateTestParFunction Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal UpdateTestParameters As Integer ) As Integer 
Declare Function SetUserEventNotify Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal EventNotify As Integer , ByVal EventCode As Integer ) As Integer 
Declare Function SetUserContactToolFunction Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal UserContactTool As Integer ) As Integer 
'  'ifdef __cplusplus
'  int WINAPI CreateTpgmWindow (HINSTANCE hInstance, BOOL bExecLoop=FALSE);
Declare Function CreateTpgmWindow Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal hInstance As Integer , ByVal bExecLoop As Integer ) As Integer 
Declare Sub RunTask Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal taskid As Integer , ByVal taskfunction As Integer )   
Declare Function BeginTask Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal dwTaskId As Integer ) As Integer 
Declare Sub EndTask Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Function BeginTest Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal dwTestNumber As Integer ) As Integer 
Declare Sub BeforeEndTest Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Sub EndTest Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Function StoreMeasure Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByVal Value As Double , ByVal Die As Short ) As Integer 
Declare Function StoreMeasureEx Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByVal Value As Double , ByVal Die As Short , ByRef CompareResult As Integer ) As Integer 
Declare Function StoreMeasureEx2 Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByVal Value As Double , ByVal Die As Short , ByRef CompareResult As Integer ) As Integer 
Declare Function StoreMeasures Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByRef Values As Double , ByRef DieList As Short ) As Integer 
Declare Function StoreMeasuresEx Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByRef Values As Double , ByRef DieList As Short , ByRef CompareResult As Integer ) As Integer 
Declare Function SaveMeasure Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByVal Value As Double , ByVal Die As Short ) As Integer 
Declare Function SaveMeasures Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByRef Values As Double , ByRef DieList As Short ) As Integer 
Declare Function StorePinsMeasures Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByRef PinList As Short , ByRef Values As Double , ByRef DieList As Short ) As Integer 
Declare Function SavePinsMeasures Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByRef PinList As Short , ByRef Values As Double , ByRef DieList As Short ) As Integer 
Declare Function StoreIdsMeasures Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByRef IdList As Short , ByRef Values As Double , ByRef DieList As Short ) As Integer 
Declare Function SaveIdsMeasures Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByRef IdList As Short , ByRef Values As Double , ByRef DieList As Short ) As Integer 
Declare Function StoreResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByVal Result As Integer , ByVal Die As Short ) As Integer 
Declare Function StoreResultEx Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByVal Result As Integer , ByVal Die As Short , ByRef CompareResult As Integer ) As Integer 
Declare Function StoreResultEx2 Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByVal Result As Integer , ByVal Die As Short , ByRef CompareResult As Integer ) As Integer 
Declare Function StoreResults Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByRef Results As Integer , ByRef DieList As Short ) As Integer 
Declare Function StoreResultsEx Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByRef Results As Integer , ByRef DieList As Short , ByRef CompareResult As Integer ) As Integer 
Declare Function SaveResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByVal Result As Integer , ByVal Die As Short ) As Integer 
Declare Function SaveResults Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByRef Results As Integer , ByRef DieList As Short ) As Integer 
Declare Function StorePinsResults Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByRef PinList As Short , ByRef Results As Integer , ByRef DieList As Short ) As Integer 
Declare Function SavePinsResults Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByRef PinList As Short , ByRef Results As Integer , ByRef DieList As Short ) As Integer 
Declare Function ClearRemarks Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer ) As Integer 
Declare Function SaveRemark Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByVal PinName As String , ByVal TestRemark As String , ByVal DieNum As Short ) As Integer 
'  BOOL WINAPI StoreDigResult (DWORD TestNum, PATTERN_ERR_PTR PatternError, WORD *DieList);
'  BOOL WINAPI SaveDigResult (DWORD TestNum, PATTERN_ERR_PTR PatternError, WORD *DieList);
'  BOOL WINAPI StorePinsDigResult (DWORD TestNum, PATTERN_ERR_PTR PatternError, WORD *PinList, WORD *DieList);
'  BOOL WINAPI SavePinsDigResult (DWORD TestNum, PATTERN_ERR_PTR PatternError, WORD *PinList, WORD *DieList);
Declare Function StoreTextInformation Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal DieNumber As Short , ByVal TextInfo As String ) As Integer 
Declare Function ExecuteTestCompare Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByRef DieList As Short ) As Integer 
Declare Function ExecuteDigTestCompare Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNum As Integer , ByRef PinList As Short , ByRef DieList As Short ) As Integer 
'  void WINAPI DisplayDigResult (PATTERN_ERR_PTR DigResult, WORD *DieList);
Declare Sub SetDieResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal DieNum As Short , ByVal BinNumber As Short )   
Declare Sub GetDieResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal DieNum As Short , ByRef Result As Short , ByRef BinNumber As Short )   
Declare Sub SetSiteResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal SiteNum As Short , ByVal BinNumber As Short )   
Declare Sub GetSiteResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal SiteNum As Short , ByRef Result As Short , ByRef BinNumber As Short )   
Declare Sub RemoveDies Lib "R3CS_FUNCTIONAL_200.DLL" ( ByRef DieList As Short )   
Declare Sub RemoveSites Lib "R3CS_FUNCTIONAL_200.DLL" ( ByRef SiteList As Short )   
Declare Sub RemoveTask Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TaskId As Integer )   
Declare Function IsRetest Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function GetDeviceNumber Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal DieNumber As Short ) As Integer 
Declare Function GetActualTaskId Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function GetActualTestNum Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Sub tcGetDiesResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByRef DieResult As Byte )   
Declare Function tcGetTestNumByPos Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestPos As Integer ) As Integer 
Declare Function tcGetTestNumMeasures Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestPos As Integer ) As Integer 
Declare Function tcGetTestIsDigResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestPos As Integer ) As Integer 
Declare Sub tcGetTestMeasData Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestPos As Integer , ByVal CntMeasure As Integer , ByRef MsType As Byte , ByRef MsPinNum As Short , ByRef MsDieNum As Short , ByRef MsValue As Double , ByRef MsResult As Byte )   
'  BYTE WINAPI tcGetTestResult (DWORD TestPos, WORD DieNumber);
Declare Function tcGetDeviceNumber Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal DieNum As Short ) As Integer 
Declare Sub tcGetDiesToStore Lib "R3CS_FUNCTIONAL_200.DLL" ( ByRef DiesToStore As Byte , ByVal Size As Integer )   
'  BYTE WINAPI tcGetTaskResult (DWORD TaskPos, WORD DieNum);
Declare Function tcGetTaskBinResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TaskPos As Integer , ByVal DieNum As Short ) As Short 
Declare Sub tcGetDieResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal DieNum As Short , ByRef Result As Byte , ByRef SwBin As Short , ByRef HwBin As Short )   
Declare Sub tcGetTestTotal Lib "R3CS_FUNCTIONAL_200.DLL" ( ByRef TotalTested As Integer , ByRef TotalPass As Integer )   
Declare Sub tcGetBinTotal Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal BinNo As Integer , ByVal DieNo As Integer , ByRef TotalBin As Integer , ByRef Yield As Double )   
Declare Sub tcGetFailSummary Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNo As Integer , ByVal DieNo As Integer , ByRef TotalExecs As Integer , ByRef TotalFail As Integer )   
'  void WINAPI tcGetLotDateTime (SYSTEMTIME *BeginDateTime, SYSTEMTIME *EndDateTime);
'  void WINAPI tcGetTestDateTime (SYSTEMTIME *BeginDateTime, SYSTEMTIME *EndDateTime);
'  void WINAPI tcGetDigResult (DWORD TestPos, PATTERN_ERR *DigResult);
Declare Function tcGetWaferNumber Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Short 
Declare Sub tcGetOcrString Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal Ocr As String )   
Declare Sub tcGetDieXY Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal DieNum As Short , ByRef XCoord As Integer , ByRef YCoord As Integer )   
Declare Function DpatInitialize Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Sub DpatDestroy Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Function DpatSaveMeasure Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function DpatComputeParameters Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function DpatAnalyse Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function DpatClearData Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Sub TpgmCtrlStartOfTest Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Sub TpgmCtrlEndOfLot Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Sub TpgmCtrlEndOfTpgm Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Sub TpgmCtrlClosePanel Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Sub TpgmLowStartTpgm Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Sub FillDynamicHeader Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal OnlyUpdateWafer As Integer )   
Declare Sub FillDynamicDevice Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Sub FillTestsDescription Lib "R3CS_FUNCTIONAL_200.DLL" ( )   
Declare Sub FillDynamicTestsResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestNumber As Integer , ByVal TestPos As Integer , ByVal DieNumber As Short )   
Declare Sub FillDynamicEventCode Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal CodeToWrite As Short )   
Declare Sub TpgmGetDieCoordinates Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal DieNumber As Integer , ByRef X As Integer , ByRef Y As Integer )   
Declare Sub TpgmGetOcrString Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal Ocr As String )   
Declare Function TpgmGetWaferNumber Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Short 
Declare Function TpgmGetDeviceNumber Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal DieNum As Short ) As Integer 
'  BYTE WINAPI TpgmGetTestResult (DWORD TestPos, WORD DieNumber);
Declare Sub TpgmGetFailSummary Lib "R3CS_FUNCTIONAL_200.DLL" ( ByVal TestPos As Integer , ByVal DieNo As Integer , ByRef TotalExecs As Integer , ByRef TotalFail As Integer )   
Declare Function IsCharacterization Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Function IsContactToolRunning Lib "R3CS_FUNCTIONAL_200.DLL" ( ) As Integer 
Declare Sub SendContactToolResult Lib "R3CS_FUNCTIONAL_200.DLL" ( ByRef PinPassCount As Short , ByRef PinShortCount As Short , ByRef PinOpenCount As Short )   
Declare Sub FuncTplanResultGet Lib "R3CS_FUNCTIONAL_200.DLL" ( ByRef pFuncTplanResult As Integer )   
'endif
'endif

End Module 
