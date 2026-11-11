# Isogen Symbol Editor XML 符号格式字段对照表

> 来源：由 `1.xml`（Symbol Editor File > Export 导出的 01HG）逆向解码，逐流 hexdump 验证。
> 解析器：临时脚本（NRBF 微流 + 尾部树结构解码）；ASC 侧规则见同目录 `ASC符号格式规则.md`。
> 官方语义参照：Hexagon Symbol Editor Help V15（Flow Arrow / Dimensioning / Mirror / Orientation / Scale Factor / Spindle SKEY / start·exit·spindle 接点）。

## 一、文件本质

```xml
<Symbol SKEY="01HG">base64(...)</Symbol>
```

- 内容 = base64 的**多个 .NET BinaryFormatter 微流拼接** + 一个自定义尾部结构（01HG 共 2048 字节）
- 每个属性值单独序列化一次：`流头 + 记录 + MessageEnd(0x0B)` 循环，共 22 个微流
- 这是 "Save to Project" / Smart Isometrics 项目库的原生格式（ASC 只是几何导出子集）

## 二、微流语法（经逐流验证）

```
[流头 17B] 00 01 00 00 00 FF FF FF FF 01 00 00 00 00 00 00 00
[可选]     06 <libId> <SKEY名>          BinaryLibrary
           0C <libId> <程序集名>        程序集绑定（Symbols, Version=7.0.0.0...）
[类定义]   04/05 <classId> <类名> <成员数> <成员名> <bt> <prim>
[实例值]   枚举类(成员=value__):  <常量2(i32)> <枚举值(i32)>
           System.Int32:           <值(i32)>
           System.Boolean:         <值(1字节)>
[结束]     0B
```

## 三、自定义尾部结构

```
0B + magic(00 50 1A A1 = 开块 / FF 50 1A A1 = 闭块) + i32 树：
根: [1, 1, 4]
 ├─ 块: [2, 0, 1000, 腿数=2, 接点数=3,
 │        腿块×N(交错计数), ...]
 │    ├─ 腿: [1, x1, y1, x2, y2, 0]     ← 实画线段，方向与 ASC 笔路一致
 │    └─ ...
 ├─ 接点块: [2, 0, 类型码, X, Y]          ← 坐标已归一化到定位点(0,0)
 └─ ...×接点数
```

## 四、01HG 全字段解码结果

### NRBF 属性流（按文件顺序）

| # | 字段 | 值 | 语义 |
|---|---|---|---|
| 1 | Symbols.eDisplayFlowArrow | 0 | Flow Arrow 属性（官方：流向箭头生成控制） |
| 2 | Symbols.eDisplayDimensions | 0 | Dimensioning 属性 |
| 3 | Symbols.eMirrorType | 0 | Mirror 属性（0=无） |
| 4 | System.Int32 | 0 | 待定（候选：ISOGEN Record，官方默认 0） |
| 5 | Symbols.eDimensionlessDir | 0 | Orientation 相关（0=None=顺管道） |
| 6 | System.Int32 | **100** | **Scale Factor = 100%** |
| 7 | System.Boolean | False | 待定标志 |
| 8 | System.Int32 | 1 | 待定（腿组数?） |
| 9 | System.Int32 | 3 | 待定（接点数 = 3 ✓） |
| 10-12 | eSymbolKeypointType + Int32×2 | 类型 **0** @ **(170,200)** | 接点1 |
| 13-15 | 同上 | 类型 **3** @ (170,200) | 接点2 |
| 16-18 | 同上 | 类型 **1** @ (170,200) | 接点3 |

### 尾部结构

| 块 | 内容 | 解读 |
|---|---|---|
| 根 | 1, 1, 4 | 版本/标志待定 |
| 头块 | 2, 0, 1000, **2**, **3** | 1000=基码；**2=腿数，3=接点数** |
| 腿1 | (0,400)→(350,400) | 上平行线（=ASC 记录1 pen2 离开段，方向一致） |
| 腿2 | (350,0)→(0,0) | 下平行线（=ASC 记录2 pen6 到达段，方向一致） |
| 接点 | 类型码 1001/1003/1002，全 @ (0,0) | 归一化坐标 |

## 五、与 ASC / 官方的交叉验证结论

1. **接点显式化**：XML 三个接点全部在 **(170,200) = ASC 首点**——证实单端支撑符号 start=exit=定位点中心；尾部再存一份归一化 (0,0)
2. **接点双编码**：NRBF 枚举值 {0, 3, 1} ↔ 尾部类型码 {1001, 1003, 1002}（按流序对应 0→1001、3→1003、1→1002）；官方语义为 start / exit / spindle，具体映射待阀门 XML 验证
3. **腿 = ASC 实画线段**：2 条腿与 ASC 解码出的两条线坐标、方向完全一致——XML 存的是解码后的几何（笔路/pen 细节不在腿块中，腿尾参恒 0）
4. **501 头部参数对上**：`100` = Scale Factor；尾参 0 = Orientation(None) 顺管道；eMirrorType=0 对应 501 参数区全 0
5. **XML 无 pen/元素号信息**：symbol/insulation/tracing/text 的区分只在 ASC 侧（元素号+pen 编码）

## 六、8 符号全库验证（2026-11-11，1.xml 补充导出 8 个支撑符号）

### eDimensionlessDir ↔ ASC 501 尾参：8/8 完美对应

| 符号 | XML eDimensionlessDir | ASC 501 尾参 | 官方 Orientation 语义 |
|---|---|---|---|
| 01HG | 0 | 0 | None 顺管道 |
| SLVE | 0 | 0 | None 顺管道 |
| ANCH | **3** | **3** | All Primary 三轴同画 |
| DUCK/GUID/HANG/SKID/SPRG | **2** | **2** | 固定方向（×5） |

### 接点编码：类型 0=START=ASC 首点，类型 1=EXIT=ASC 收笔点（8/8 全部实锤）

| 符号 | 接点枚举 | START 坐标 | EXIT 坐标 | ASC 首点/收笔点 |
|---|---|---|---|---|
| 01HG | 0, **3**, 1 | (170,200) | **(170,200)** | 重合（笔回家）✓ |
| ANCH | 0, 1 | (2000,250) | (4000,250) | ✓✓ |
| DUCK | 0, 1 | (0,300) | (1100,300) | ✓✓ |
| GUID | 0, 1 | (400,600) | (3300,600) | ✓ |
| HANG | 0, 1 | (0,500) | (2500,500) | ✓✓ |
| SKID | 0, 1 | (0,500) | (2500,500) | ✓ |
| SLVE | 0, 1 | (500,300) | **(500,300)** | 重合 ✓ |
| SPRG | 0, 1 | (0,500) | (3350,500) | ✓✓ |

**结论**：ASC 格式用"首点=START、收笔点=EXIT"编码双接点；01HG/SLVE 等 start=exit 重合于中心。
01HG 独有第三个接点（类型 3 @ 中心，候选 tap point）；DUCK 独有 eMirrorType=1（其余全 0）。

### 尾部腿块（部分校准）

- 01HG：2 条腿 (0,400)→(350,400)、(350,0)→(0,0)——与 ASC 实画线段完全一致
- ANCH：轴线 (0,250)→(4000,250) 精确命中；箭头腿块有 4 字节错位（块内疑含点数/标志字段）
- GUID：扇形排线 (800,600)→(1800,~100 步进) 清晰可见，共 21 条腿
- 多段折线符号的腿块内部语法（点数/标志/0xFFFFFFFF 哨兵）待下一轮逐块校准

## 七、阀门+阀杆验证（2026-11-11，22.xml / VT.asc / 01sp.asc / 05SP.asc）

### 接点类型码表（VT** 锁定）

| NRBF 枚举 | 尾部码 | 语义 | VT** 坐标 |
|---|---|---|---|
| 0 | 1001 | **START** | (0,200) = ASC 首点 ✓ |
| 3 | 1003 | **SPINDLE**（阀杆接点） | (400,200) 中心 |
| 1 | 1002 | **EXIT** | (800,200) = ASC 收笔点 ✓ |

文件中接点顺序恒为 start→spindle→exit。01HG 也带类型3中心点（支撑带 spindle 码，语义待官方解释）。

### 阀杆符号（T 形：横杆+竖柄）

- **01SP**：START (0,200)、EXIT (400,200)；腿 (0,200)→(400,200) + (400,400)→(400,0)——与 ASC 笔路一致
- **05SP**：START=EXIT=(240,192)；腿 (0,192)→(640,192) + (640,400)→(640,0)
- 组合机制：`501 VT 05SP` 只存阀体几何+SKEY 绑定；阀柄由独立 `501 05SP` 定义，渲染时对齐阀体 spindle 接点 (400,200) 叠加

### 其他确认

- `Int32[9]` = 接点数（支撑=2，01HG/VT**=3）
- VT** 独有 `eDisplayFlowArrow=2`（官方 Flow Arrow 属性；支撑全 0）
- DUCK/SKID/SPRG/SLVE 的尾部腿块仍含非坐标字段错位；单线段符号（01HG/01SP/05SP）干净对齐

## 八、待定项与验证建议

- ~~接点类型码与 start/exit/spindle 对应~~ ✓ 已由 VT** 锁定（0/1/3 ↔ 1001/1002/1003）
- `Int32[8]`（恒 1）与 `Boolean(False)` 的准确语义
- 尾部根标量 [1, 1, 4] 的含义；类型 2 接点码（候选 tap point）尚未出现
- 多段折线符号腿块内部字段（点数/标志/0xFFFFFFFF 哨兵）校准
