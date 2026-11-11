# Isogen SKEY 总表（全量机器提取）

> 来源：`C:\Win32app\Smart3D\DrawingsIsometric\Alias\Core Components\SystemMaterialData.xml`（Smart3D 自带 SystemMaterialData.xml，A_VERSION 1.4）
> 组件定义 180 条 · 唯一 SKEY 455 个 · 内嵌符号库 363 个（解码成功 363）
> 接点类型：0=START 1=EXIT 3=SPINDLE（见 XML符号格式字段对照表.md）

## SKEY 组合规则与端部形式（End Preparations）

**SKEY = 前两位基形 + 后两位端部形式**（官方规则），如 `VTPL` = `VT**`（阀门基形）+ `PL`（plain end）。

渲染组合机制（端部符号贴在零件接点处）：
1. 组件基形符号（如 VT**）按首点/收笔点对齐组件 END-POINT 绘制
2. 端部形式符号贴到组件的 START/EXIT 接点上——**端部符号自身也是两接点小符号**
   （如 PL: START(0,150)↔EXIT(200,150)；PF: (0,200)↔(240,200)），其 START 贴组件、EXIT 沿管轴朝外
3. 管道接到端部符号的 EXIT——即"阀门两端先接到 PL，然后 PL 再接管"

另有双端组合 SKEY（两端不同端部形式）：`FLFL`、`FLSC`、`FLSW`、`COSC`、`COSW`、`KABW` 等。

**66 条端部形式定义**（A_ENDPREP 区，格式 `BW--`，含 Isogen/Trimmable 标志）：
BW=Butt Weld(可修切)、FL=Flanged、SC=Screwed、SW=Socket Weld、PL=Plain End、CP=Compression、
LN=Liner Nut、LC=Liner/Clamp、LR=Reducing Liner/Nut、MP=Male Part、PF=Push Fit、GL=Glued、
CL=Clamped、FA=Flared、BS/SB=Ball&Socket、GF=Gland 等（完整清单见 skeys.json / 源文件）

**25 个端部准备符号**（有独立几何+双接点）：
PL RP SP SR SB OB PF PV RD OP PR DR WW WM WS WF WO K7 MP LC LN LR BM BP XX


## AMAT

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Additional Material | ADDITIONAL-ITEM | Additional Material | Dimensionless | AMAT BOLT FITT FLAN GASK INST MISC PIPE SUPP VALV |

## BOLT

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Bolt | BOLT | Bolt | Dimensionless | BOLT |

## EQPT

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Equipment | EQUIPMENT | Equipment | NonStandard | XB XC XD |
| Nozzle | NOZZLE | Nozzle | NonStandard | BW CP FL LC LN LR MP PL SC SW PF GL FA CL BS LJ |

## FITT

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Fixed Pipe | PIPE-FIXED | Fixed length Pipe | Inline | FP** |
| Backing Nut | NUT | Backing Nut (Hygienic) | Dimensionless | BNUT |
| Clamp | CLAMP | Clamp | Dimensionless | CLMP CLVT CLVR CLCS CLGY |
| Inline Clamp | CLAMP | Clamp | Inline | CLMP CLVT CLVR CLCS CLGY |
| Connector | CONNECTOR | Connector | Inline | LC** LCEX LN** LNEX LR** LREX MP** MPEX LVBW |

## FLAN

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Slip on Flange | FLANGE | Flange | Inline | FLSO FLSJ FOSO |
| Weld Neck Flange | FLANGE | Flange | Inline | FLWN FOWN |
| Backing Flange | FLANGE | Backing Flange | Inline | FLFL FLLB |
| Lapjoint StubEnd | LAPJOINT-STUBEND | Lapjoint StubEnd | Inline | FLSE |
| Lapjoint Ring | LAPJOINT-RING | Lapjoint Ring | Inline | FLRG |
| Lapjoint StubEnd Ring Combined | FLANGE | Lapjoint StubEnd Ring Combined | Inline | FBSE |
| Flange | FLANGE | Flange | Inline | FLSC FLSW PUD TWEL FLSF FLSM FLGF FLGM FLGL FLPF FLFF FLFR |
| Ball/Socket Flange | FLANGE | Flange | Inline | FLBS FLSB |
| Concentric Reducing Flange | FLANGE-REDUCING-CONCENTRIC | Concentric Reducing Flange | Inline | FLRC |
| Eccentric Reducing Flange | FLANGE-REDUCING-ECCENTRIC | Eccentric Reducing Flange | Offset_2D | FLRE |
| Weld Neck Jacket Flange | FLANGE | Jacket Flange | NonStandard | JFWN |
| Slip on Jacket Flange | FLANGE | Jacket Flange | NonStandard | JFSO |
| Concentric Reducing Jacket Flange | FLANGE-REDUCING-CONCENTRIC | Concentric Reducing Jacket Flange | NonStandard | JFRC |
| Eccentric Reducing Jacket Flange | FLANGE-REDUCING-ECCENTRIC | Eccentric Reducing Jacket Flange | NonStandard | JFRE |
| Hygienic Fittings | MISC-HYGIENIC | Miscellaneous Hygienic fittings | InlineWithCentre | ADMF ADMM DVP BBC BBE BM BTP |

## Fittings

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Elbow 90 | ELBOW | 90 Deg Elbow | Angle_90 | EL** ER** EBSC |
| Elbow 45 | ELBOW | 45 Deg Elbow | Angle_45 | EL** ER** EBSC |
| Elbow (Non standard angle) | ELBOW | Elbow - non standard angle | NonStandard | EL** ER** EBSC |
| Concentric Reducer | REDUCER-CONCENTRIC | Concentric Reducer | Inline | RC** CP** CS** RB** RF** RN** |
| Eccentric Reducer | REDUCER-ECCENTRIC | Eccentric Reducer | Offset_2D | RE** EP** ES** |
| Fitting Tee | TEE | Fitting Tee | Branch | TE** TS** |
| Weldolet | OLET | Weldolet Olet | Inline | WT** |
| Sockolet | OLET | Sockolet olet | Inline | SK** |
| Latrolet | OLET | Latrolet olet | Inline | LA** |
| Nipolet | OLET | Nipolet olet | Inline | NI** K7 |
| Sweepolet | OLET | Sweepolet olet | Inline | SW** |
| Thredolet | OLET | Thredolet olet | Inline | TH** |
| Nippoflange | OLET | Nippoflange olet | Inline | WOFL |
| Half Coupling | OLET | Half coupling olet | Inline | HC** |
| Fitting Cross | CROSS | Cross  (fitting type) | Cross | CR** CY** |
| Teed Concentric Reducer | REDUCER-CONCENTRIC-TEED | Teed Concentric Reducer | NonStandard | CT** CX** CZ** |
| Teed Eccentric Reducer | REDUCER-ECCENTRIC-TEED | Eccentric Reducer with connection | NonStandard | EZ** OT** EXBW EXFL |
| Teed Elbow | ELBOW-TEED | Elbow - (90 degrees) with a connection | NonStandard | ET** |
| Return Elbow | ELBOW | Elbow - 180 degree return ( U elbow ) | Return_2D | EU** |
| Pulled Bend Fabricated | BEND | Pulled Bend | Angle_90 | BE** |
| Mitred Bend | BEND | Mitred Bend | Angle_90 | L1** L2** L3** L4** L5** L6** L7** L8** L9** L@** MI** |
| Return Bend Fabricated | BEND | Pulled 180 deg bend | Return_2D | BUFL |
| Teed Bend | BEND-TEED | Bend - with a connection | NonStandard | BTBS BTFL TBBW TBGF MTBW MTFL T1BW T1FL T1PL T2BW T2FL T2PL T3BW T3FL T3PL T4BW T4FL T4PL T5BW T5FL T5PL T6BW T6FL T6PL T7BW T7FL T7PL T8BW T8FL T8PL T9BW T9FL T9PL T@BW T@FL T@PL TB TB1D TB2D TB3D TB4D TB5D TB6D TB7D TB8D TB9D |
| Tap | TAP-CONNECTION | Tapping point | Branch |  |
| Fitting Y-Piece | Y-PIECE-FITTING | Y-Type Tee | NonStandard | TY** |
| Instrument Tee | INSTRUMENT-TEE | Instrument Tee | Branch | IT** |
| Jacket Tapping Point | JACKET-CONNECTION | Jacket tapping point | Branch |  |

## GASK

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Gasket | GASKET | Gaskets | Inline | GASK |

## INFO

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Pipeline Drawing Sequence Number | PIPELINE-DRAWING-SEQUENCE-NUMBER | Pipeline Drawing Sequence Number | None |  |
| Pipeline Drawing Number | PIPELINE-DRAWING-SEQUENCE-NUMBER | Pipeline Drawing Sequence Number | None |  |
| Highest Part Number | HIGHEST-PART-NUMBER | Highest part number used in pipeline | None |  |
| Highest Assembly Number | HIGHEST-ASSEMBLY-NUMBER | Highest Assembly number used in pipeline | None |  |
| Highest Spool Number | HIGHEST-SPOOL-NUMBER | Highest Spool number used in pipeline | None |  |
| Highest Unique Identifier | HIGHEST-UNIQUE-IDENTIFIER | Highest unique identifier used in pipeline | None |  |
| Highest Weld Number | HIGHEST-WELD-NUMBER | Highest weld number used in pipeline | None |  |
| Highest Support Weld Number | HIGHEST-SUPPORT-WELD-NUMBER | Highest Support Weld number used in pipeline | None |  |
| Flow Arrow | FLOW-ARROW | Flow Indicator | Dimensionless | FLOW |
| Location Point | LOCATION-POINT | Location Point | Dimensionless | LOPT |
| Nozzle | NOZZLE | Nozzle | End | NOZZ |
| Message | MESSAGE | Message | None |  |
| Message Circle | MESSAGE-CIRCLE | Message in a circle | None |  |
| Message Diamond | MESSAGE-DIAMOND | Message in a diamond | None |  |
| Message Double Circle | MESSAGE-DOUBLE-CIRCLE | Message in a double circle | None |  |
| Message Ellipse | MESSAGE-ELLIPSE | Message in an ellipse | None |  |
| Message Pointed | MESSAGE-POINTED | Message in a pointed ended box | None |  |
| Message Round | MESSAGE-ROUND | Message in a round ended box | None |  |
| Message Square | MESSAGE-SQUARE | Message in a square ended box | None |  |
| Message Triangle | MESSAGE-TRIANGLE | Message in a triangle | None |  |
| Break in Point | BIP-IDENTIFIER | Break in point identifier | Dimensionless |  |
| Induction Bend Start | INDUCTION-START | Induction bend start point | Dimensionless |  |
| Induction Bend End | INDUCTION-END | Induction bend end point | Dimensionless |  |
| Isometric Split Point | ISO-SPLIT-POINT | Isometric split point | None |  |
| System Split Point | SYSTEM-SPLIT | System Isometric split point | None |  |
| Cut Piece Allowance | CUT-PIECE-ALLOWANCE | Cutting allowance for pipe | None |  |
| Weld Number | REPEAT-WELD-IDENTIFIER | Repeat Weld Identifier | None |  |
| Spool Identifier | SPOOL-IDENTIFIER | Spool Identifier | None |  |
| Detail Sketch Identifier | DETAIL-SKETCH-IDENTIFIER | Detail sketch / information note | None |  |
| Client Drawing Identifier | CLIENT-DRAWING-IDENTIFIER | Client Drawing Identifier | None |  |
| Bypass Closure Point | BYPASS-CLOSURE-POINT | Bypass closure point | None |  |
| End Connection Pipeline | END-CONNECTION-PIPELINE | End Connection Pipeline | None |  |
| End Connection Equipment | END-CONNECTION-EQUIPMENT | End Connection Equipment | None | NZFS NZFE NZWS NZWE |
| End Position Null | END-POSITION-NULL | End Position Null | None |  |
| End Position Open | END-POSITION-OPEN | End Position Open | None |  |
| End Position Closed | END-POSITION-CLOSED | End Position Closed | None |  |
| End Position Vent | END-POSITION-VENT | End Position Vent | None |  |
| End Position Drain | END-POSITION-DRAIN | End Position Drain | None |  |
| End Connection Jacket | END-CONNECTION-JACKET | End Connection Jacket | None |  |
| End Connection Core | END-CONNECTION-CORE | End Connection Core | None |  |
| Thickness Measurement Location | TML | Thickness Measurement Location | None | THML |
| End Preparation Types | END-PREP | End Preparation Types | None | BW CP FL LC LN LR MP PL SC SW PF GL FA CL BS LJ |

## INST

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Control Valve | INSTRUMENT | Instrument | InlineWithCentre | CV** HV** IG** MV** RV** SV** ZV** |
| Angle Control Valve | INSTRUMENT-ANGLE | Angle Control Valve | Angle_90 | CA** HA** IA** MA** RA** SA** XA** |
| 3 Way Control Valve | INSTRUMENT-3WAY | 3 Way Control Valve | Branch | C3** H3** M3** S3** |
| 4 Way Control Valve | INSTRUMENT-4WAY | 4 Way Control Valve | Cross | C4** H4** M4** S4** |
| Instrument | INSTRUMENT | Instrument | InlineWithCentre | II** DR MONO PR |
| Instrument Dial | INSTRUMENT-DIAL | Instrument Dial | End | ID** |
| Instrument - External | INSTRUMENT-EXTERNAL | Instrument - External | Oriented | IEPL |
| Instrument - Offset | INSTRUMENT-OFFSET | Instrument - Offset | Offset_2D | IO** |
| Instrument - Return | INSTRUMENT-RETURN | Instrument - Return (2D) | Return_2D | IR** |
| Instrument - Angle | INSTRUMENT-ANGLE | General Instrument - Angle | Angle_90 | IA** |
| Orifice Plate | INSTRUMENT | Instrument | InlineWithCentre | OP |
| Instrument Balloon | INSTRUMENT-BALLOON | Instrument Balloon | End | IB** |

## Instruments

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Restrictor Plate | MISC-COMPONENT | Restrictor Plate | InlineWithCentre | RP |

## Miscellaneous

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Cap | CAP | Cap | End | KA** |
| Jacket Cap | CAP | Jacket Cap | End | KJ** |
| Variable Length Pipe Block | PIPE-BLOCK-VARIABLE | Variable length pipe block | Inline | PV PVIJ |
| Filter | FILTER | Inline Filter | Inline | FI** YS** |
| Angle Filter | FILTER-ANGLE | Angle Filter | Angle_90 | FA** |
| Offset Filter | FILTER-OFFSET | Offset Filter | Offset_2D | FO** |
| Return Filter | FILTER-RETURN | Return Filter | Return_2D | FR** |
| Miscellaneous Component | MISC-COMPONENT | Miscellaneous | InlineWithCentre | CH** EX** FT** FX** NC** TU** |
| Plug | MISC-COMPONENT | Miscellaneous | End | PL |
| Spectacle Blind | MISC-COMPONENT | Miscellaneous | InlineWithCentre | SB OB |
| Sight Glass | MISC-COMPONENT | Miscellaneous | InlineWithCentre | SG** |
| Slip Plate | MISC-COMPONENT | Miscellaneous | InlineWithCentre | SP |
| Slip Ring | MISC-COMPONENT | Miscellaneous | InlineWithCentre | SR |
| Angle | MISC-COMPONENT-ANGLE | Miscellaneous | Angle_90 | BA** |
| Offset | MISC-COMPONENT-OFFSET | Miscellaneous Offsets | Offset_2D | BO** |
| Return | MISC-COMPONENT-RETURN | Miscellaneous Returns | Return_2D | BR** |
| Coupling | COUPLING | Coupling | Inline | WTBW NISC NIPL NRSC NBSC COCP COPF COGL COCL COFA COVR COGY CSCP COSC COSW SWBW SKSW LABW LASC LASW |
| Elbolet | ELBOLET | Elbolet | Inline | CE** |
| Union | UNION | Union | Inline | UN** |
| Reinforcement Pad | REINFORCEMENT-PAD | Reinforcement Pad | Dimensionless | RPAD |
| Gap | GAP | Gap | Inline | GAP |
| Penetration Plate | PENETRATION-PLATE | Penetration Plate | Dimensionless | CRPP |
| Pipe Break | PIPE-BREAK | Pipe Break | Dimensionless | PBRK |

## PIPE

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Pipe | PIPE | Pipe | Inline | PIPE |
| Set On Tee | TEE-SET-ON | Set on Tee | Inline | TESO TERF TPUL TTSO TTRF TOSO TORF |
| Tee Stub | TEE-STUB | Tee Stub | Inline | TSSO TSRF |
| Set On Cross | CROSS-SET-ON | Cross (set on type) | Inline | CRSO CRRF CYSO |
| Pulled Bend | BEND | Pulled Bend | Angle_90 | PB PB+D PB1D PB2D PB3D PB4D PB5D PB6D PB7D PB8D PB9D |
| Return Bend | BEND | Pulled 180 deg bend | Return_2D | BU BU+D BU1D BU2D BU3D BU4D BU5D BU6D BU7D BU8D BU9D |
| Fabricated Y-Piece | Y-PIECE-FABRICATED | Y-Type Fabricated Tee | NonStandard | TYSO |

## SUPP

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Support | SUPPORT | Pipe Support | Dimensionless | 01HG  ANCH DUCK GUID SKID SPRG HANG SLVE |

## SYMB

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Reference Dimension | REFERENCE-DIMENSION | Reference Dimension | Dimensionless | BLD BLD1 BLD2 BLD3 BLD4 EQU EQU1 EQU2 EQU3 EQU4 FLR FLR1 FLR2 FLR3 FLR4 GRD GRD1 GRD2 GRD3 GRD4 HST HST1 HST2 HST3 HST4 PIP PIP1 PIP2 PIP3 PIP4 VST VST1 VST2 VST3 VST4 WAL WAL1 WAL2 WAL3 WAL4 XXX XXX1 XXX2 XXX3 XXX4 |
| Floor Symbol | FLOOR-SYMBOL | Floor Symbol | Dimensionless | FLOR |
| Wall Symbol | WALL-SYMBOL | Wall Symbol | Dimensionless | WALL |
| Insulation Symbol | INSULATION-SYMBOL | Insulation symbol | Dimensionless | INPP |

## VALV

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Ball Valve | VALVE | Ball Valve | InlineWithCentre | VB** |
| Check Valve | VALVE | Check Valve | InlineWithCentre | VC** CK** |
| Globe Valve | VALVE | Globe Valve | InlineWithCentre | VG** |
| Gate Valve | VALVE | Gate Valve | InlineWithCentre | VT** |
| Butterfly Valve | VALVE | Butterfly Valve | InlineWithCentre | VY** ZB** |
| Diaphragm Valve | VALVE | Standard Gate Valve | InlineWithCentre | VD** |
| Cock Valve | VALVE | Cock Valve | InlineWithCentre | KV** VK** |
| Needle Valve | VALVE | Needle Valve | InlineWithCentre | VN** |
| Plug Valve | VALVE | Plug Valve | InlineWithCentre | VP** |
| Relief Valve | VALVE | Relief Valve | InlineWithCentre | VR** |
| Slide Valve | VALVE | Slide Valve | InlineWithCentre | VS** |
| Basic Valve | VALVE | Basic Valve | InlineWithCentre | VV** ZG** |
| Pressure Reducing Valve | VALVE | Pressure Reducing Valve | InlineWithCentre | VX** |
| Pressure Relief Valve | VALVE | Pressure Relief Valve | InlineWithCentre | VZ** |
| Nonreturn Valve | VALVE | Non Return Valve | InlineWithCentre | NV** |
| Angle Valve | VALVE-ANGLE | Angle Valve | Angle_90 | AR** AV** AX** 2D** 2Z** |
| 3 Way Valve | VALVE-3WAY | 3 Way Valve | Branch | V3** K3** 3D** 3Z** |
| 4 Way Valve | VALVE-4WAY | 4 Way Valve | Cross | V4** 4D** 4Z** |
| Multi way valve | VALVE-MULTIWAY | Multi way valve | NonStandard | MD** MZ** |
| Two Level 3 Port | VALVE-MULTIWAY | Two Level 3 Port Valve | NonStandard | MD** MZ** |
| Two Level 4 Port | VALVE-MULTIWAY | Two Level 4 Port Valve | NonStandard | MD** MZ** |
| Blind Flange | FLANGE-BLIND | Blind (Blank) Flange | End | FLBL |
| Safety Disc | SAFETY-DISC | Rupture Disc | InlineWithCentre | RD |
| Fixed Length Pipe Block | PIPE-BLOCK-FIXED | Fixed length pipe block | Inline | PF PFIJ |
| Trap | TRAP | In Line Trap | Inline | TI** |
| Angle Trap | TRAP-ANGLE | Angle Trap | Angle_90 | TA** |
| Offset Trap | TRAP-OFFSET | Offset Trap | Offset_2D | TO** |
| Return Trap | TRAP-RETURN | Return Trap | Return_2D | TR** |
| Multi Port Component | MULTI-PORT-COMPONENT | Multi Port Component | Inline | XV** XF** |

## WELD

| 组件类型 | PCFTYPE | 描述 | 配置 | SKEY |
|---|---|---|---|---|
| Mitre Weld | WELD | Mitred Weld | NonStandard | WM WMF WMFT WMT WMS WMO WMD WMSD WMOD |
| Workshop Weld | WELD | Workshop Weld | Inline | WW WWA WWD WWST |
| Workshop Jacket Weld | WELD | Workshop Jacket Weld | Inline | WWJ |
| Site Weld | WELD | Site/Field/Erection Weld | Inline | SW2 WF WFST WFT WFD WS WSST WST WSD WSSP WSSR XX XXD |
| Site Jacket Weld | WELD | Site Jacket Weld | Inline | WSJ |
| Offshore Weld | WELD | Offshore Weld | Inline | WO WOF WOD WOFT WOT WOSR WOST WVST WOFD |
| Offshore Jacket Weld | WELD | Offshore Jacket Weld | Inline | WOJ |
| Support Weld | WELD | Support Weld | Inline | ZSP1 ZSP2 ZSP3 ZSP4 ZSP5 ZSP6 ZSP7 ZSP8 ZSP9 ZTN1 ZTN2 ZTR1 ZTR2 ZUD1 |
| End Plate Weld | WELD | End Plate Weld | Inline | ZEP1 ZEP2 |

## 符号库统计

- Orientation 分布: {'0': 356, '1': 1, '3': 1, '2': 5}
- 含 SPINDLE 接点的符号 62 个: 01HG 2D** 2Z** 3D** 3Z** 4D** 4Z** AR** AV** AX** BO** BR** C3** C4** CA** CK** CV** FR** H3** H4** HA** HV** IG** IO** IR** IT** K3** KV** M3** M4** MA** MD** MV** MZ** RA** RV** S3** S4** SA** SV** TR** V3** V4** VB** VC** VD** VG** VK** VN** VP** VR** VS** VT** VV** VX** VY** XA** YLRG YMED YSML ZB** ZG**