# -*- coding: utf-8 -*-
"""
microgrid_SystemV1.2.py —— 园区微电网风光储协调优化配置 · 仿真验证平台
================================================================================
把《园区微电网风光储协调优化配置研究》全篇论文的计算链路整合为单一可交互程序，
用于验证作者的各类能耗调配设想、模拟模型运行、给出计算结果，并显示/产出图像与数据。

--------------------------------------------------------------------------------
 一、功能总览（与题目 3.3 基础要求、3.4 拓展要求一一对应）
--------------------------------------------------------------------------------
  ① 数据检查与曲线   校验内置/载入数据（时段数、粒度、缺失、量纲、越界、一致性）；
                     绘制三园区典型日负荷 / 风光出力 / 净负荷 / 月度热力图（3.3-1、3.3-2）
  ② 无储能基线       确定性算术，给出购电、弃电、总成本、单位成本（3.3-4）
  ③ 储能运行模拟     给定 50 kW/100 kWh，人工规则 vs LP 最优 vs 无储能（3.3-5、3.4-1）
  ④ 最优运行策略     LP 最优逐时调度与 SOC 轨迹，含调度明细表（3.4-1）
  ⑤ 储能容量优化     (P,E) 网格 + 迭代细化，论证 50 kW/100 kWh 是否最优（3.4-2）
  ⑥ 联合运营分析     三园区统一运营 vs 独立运营，共六情景对比（3.4-3）
  ⑦ 敏感性分析       储能单价 / 购电电价 / 充放电效率 / 运行寿命（3.4-6）
  ⑧ 结果自检         物理可行性 + 双求解器交叉验证（单次结果）
  ⑨ 一键回归自检     16 项内置回归，固定用内置模板与题目默认参数
  导出报表 / 停止计算 / 清空

  长任务（⑤⑥⑦⑨）：状态栏显示进度百分比，可随时点「停止计算」中止。

--------------------------------------------------------------------------------
 二、数据来源（V1.1 起）
--------------------------------------------------------------------------------
  1) 默认使用【代码内置数据模板】：三个附件的数据已固化在 §3.0 的 _EMBEDDED 中，
     因此本程序不依赖任何外部文件，可单独拷走运行（真正的单文件）。
  2) 「载入 Excel 数据…」可用自备附件替换（按文件名自动识别附件1/2/3，
     缺哪个就用内置模板补齐那部分）。
  3) 「恢复内置数据」切回内置模板。
  4) 界面可直接调整【负荷缩放%】【风电装机】【光伏装机】，
     kW 出力 = 标幺值 × 装机，改完重跑 ①~⑦ 即可对比。

--------------------------------------------------------------------------------
 三、可调参数（界面左侧面板，改动即时生效）
--------------------------------------------------------------------------------
  分析对象（A/B/C/联合）、数据时段（典型日 或 1~12 月）、负荷缩放、三园区装机、
  购电电价、风电/光伏使用成本、是否允许售电、是否启用分时电价（峰/平/谷）、
  储能功率 P、容量 E、功率/容量单价、充放电效率、SOC 上下限、运行寿命、年运行天数、
  是否启用 MILP 互斥约束、求解时限、容量优化网格范围与细化轮数、敏感性网格精度。

--------------------------------------------------------------------------------
 四、性能设计（V1.1 起）
--------------------------------------------------------------------------------
  · 不允许售电时（默认）自动剔除外送变量，LP 规模由 8T+1 降到 6T+1。
  · 批量网格搜索（容量优化/联合运营/敏感性）关闭每格的 PuLP 复算，避免反复
    启动 CBC 子进程；单次容量优化仍保留 PuLP 交叉验证。
  · 敏感性分析提供 粗/中/细 三档网格；若未安装 SciPy（网格点需逐点启动 CBC 子进程，
    慢约两个数量级），会自动放粗网格并给出安装提示。

--------------------------------------------------------------------------------
 五、求解器与自动降级（V1.2 起，重要）
--------------------------------------------------------------------------------
  PuLP 自带的 CBC 是可执行文件，PuLP 会把 cbc.exe 释放到 %TEMP% 再启动；在部分
  Windows 机器上（缺 MSVC 运行库 / 杀毒软件拦截 %TEMP% / 临时目录受限）该子进程
  无法启动，会抛：
      PulpSolverError: Pulp: cannot execute ...\\solverdir\\cbc\\win\\64\\cbc.exe
  导致 ③④⑤⑦ 全部中断。

  V1.2 起把 SciPy/HiGHS 作为首选引擎（进程内、随 wheel 分发、不依赖外部 exe、
  也不需要 MSVC 运行库），并做三层自动降级：
        PuLP/CBC  →  SciPy/HiGHS  →  纯 Python 规则策略
  任何一层不可用或调用失败都会自动切换并在日志中说明，绝不中断计算。
  启动时自动做一次求解器自检（真的启动一次 CBC 试算），据此：
        · 显示当前主求解引擎；
        · CBC 不可用时自动禁用「MILP 硬约束」并解释原因；
        · 无 SciPy 时给出 pip 安装建议；
        · 若连 SciPy 也没有，网格搜索退化为规则策略
          （「单一电价 + 不允许售电」下规则策略即精确最优，见论文 5.3 节；
            分时电价 / 允许售电场景下为近似值，日志会明确提示）。

  命令行自检（无需界面，可验证打包后的 exe）：
        microgrid_SystemV1.2.exe --selftest
     结论打印到控制台，同时写入同目录 selftest_report.txt。

--------------------------------------------------------------------------------
 六、环境依赖
--------------------------------------------------------------------------------
  必需：pyqt5、numpy、matplotlib（读 Excel 还需 openpyxl）
  强烈推荐：scipy —— 进程内求解器，最可靠，网格搜索快（未装则显著变慢）
  可选：pulp —— 仅在需要 MILP 硬约束或 CBC 交叉验证时才用到
  pip install pyqt5 numpy matplotlib openpyxl scipy pulp

--------------------------------------------------------------------------------
 七、修订记录
--------------------------------------------------------------------------------
  V1.2  ① 修复：部分 Windows 机器上 PulpSolverError（CBC 无法执行）导致
            ③④⑤⑦ 全部中断的问题。改为以 SciPy/HiGHS 为首选求解引擎，
            PuLP 调用全部加保护并三层自动降级；新增启动求解器自检，
            CBC 不可用时自动禁用 MILP 并给出修复建议（装 VC++ 运行库 / 杀软放行 %TEMP%）。
        ② 新增 --selftest 命令行自检模式（含离屏 GUI 栈自检），共 17 项；
            便于在无界面环境下验证打包后的 exe 是否环境齐备、数值正确。
        ③ 无 CBC 时交叉验证改用「HiGHS ↔ 规则策略」互校（原为 HiGHS ↔ PuLP）。
        ④ 默认参数下所有数值结果与 V1.1 完全一致。
        ⑤ 打包版 exe 采用 Calc.ico 作为程序图标（exe 文件图标由 --icon 指定，
            窗口/任务栏图标由 icon_path() 在运行时加载）。
  V1.1  ① 新增：把三个附件数据内置为 _EMBEDDED 默认模板，程序可脱离外部文件独立运行；
            新增「载入 Excel 数据…」「恢复内置数据」；数据层改为"标幺值 + 装机"结构，
            界面可直接调负荷缩放与三园区装机。
        ② 修复：⑦ 敏感性分析在默认网格下耗时 >10 分钟且界面无反馈的问题。原因有三：
            (a) 敏感性/联合运营的网格步长被硬编码，未随界面设置变化；
            (b) 每个网格点都调用一次 PuLP 复算，反复启动 CBC 子进程；
            (c) 全过程只在结束时刷新界面。修复后默认精度约 30 秒完成
            （实测 28.2 s / 78 个配置）。
        ③ 新增：长任务进度百分比显示 + 「停止计算」按钮（网格搜索逐点响应中止）。
        ④ 优化：不允许售电时 LP 变量由 8T+1 精简为 6T+1。默认参数下所有结果与 V1.0
            完全一致（基线 0.7701 / 0.6577 / 0.6383 元/kWh）。
  V1.0  首版：整合 mb_core / mb_model / mb_stage2 全部计算链路为单文件 GUI；
        新增分时电价与允许售电两个"设想验证"开关；新增结果自检与报表导出；
        默认参数严格复现论文数值（基线 0.7701 / 0.6577 / 0.6383 元/kWh）。
================================================================================
"""
from __future__ import annotations

# ---- 必须在 numpy / 求解器之前设置：抑制数学库多线程竞争导致的进程级崩溃 ----
import os

os.environ.setdefault("MKL_NUM_THREADS", "1")
os.environ.setdefault("OMP_NUM_THREADS", "1")
os.environ.setdefault("NUMEXPR_NUM_THREADS", "1")
os.environ.setdefault("OPENBLAS_NUM_THREADS", "1")
os.environ.setdefault("QT_AUTO_SCREEN_SCALE_FACTOR", "1")

import csv
import datetime as _dt
import math
import os.path as _osp
import sys
import threading
import time
import traceback
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

import numpy as np

APP_NAME = "园区微电网风光储协调优化配置仿真平台"
APP_VERSION = "V1.2"
APP_FILE = "microgrid_SystemV1.2.py"

PARKS = ("A", "B", "C")
PCOLOR = {"A": "#C0392B", "B": "#1F6FB4", "C": "#1E8449", "JOINT": "#7D3C98"}
HOURS_VALLEY = {0, 1, 2, 3, 4, 5, 23}
HOURS_PEAK = {8, 9, 10, 11, 18, 19, 20, 21}

# ------------------------------------------------------------------ 绘图
try:
    import matplotlib

    matplotlib.use("Qt5Agg")
    from matplotlib import rcParams

    rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans", "sans-serif"]
    rcParams["axes.unicode_minus"] = False
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
    from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavToolbar
    from matplotlib.figure import Figure

    HAS_MPL = True
except Exception:  # pragma: no cover
    HAS_MPL = False

# ------------------------------------------------------------------ 求解器
try:
    import pulp as pl

    HAS_PULP = True
    PULP_VER = getattr(pl, "__version__", "?")
except Exception:  # pragma: no cover
    pl = None
    HAS_PULP = False
    PULP_VER = "-"

try:
    from scipy.optimize import linprog as _linprog

    HAS_SCIPY = True
except Exception:  # pragma: no cover
    _linprog = None
    HAS_SCIPY = False

try:
    import openpyxl

    HAS_XL = True
except Exception:  # pragma: no cover
    openpyxl = None
    HAS_XL = False

TOL = 1e-6


# ==============================================================================
#  §0  中止/取消机制（网格搜索等长任务可被界面按钮中止）
# ==============================================================================
class Cancelled(Exception):
    """用户主动中止计算。"""


_CANCEL = threading.Event()


def request_cancel() -> None:
    _CANCEL.set()


def clear_cancel() -> None:
    _CANCEL.clear()


def check_cancel() -> None:
    if _CANCEL.is_set():
        raise Cancelled("计算已被用户中止")


# ==============================================================================
#  §0.5  求解器能力探测与自动降级（V1.2）
#  ----------------------------------------------------------------------------
#  背景：PuLP 自带的 CBC 是一个可执行文件，PuLP 运行时会先把它释放到 %TEMP%
#  再启动。在部分 Windows 机器上（缺 MSVC 运行库 / 杀毒软件拦截 %TEMP% /
#  临时目录受限）这个子进程无法启动，PuLP 抛 PulpSolverError: cannot execute…，
#  导致 ③④⑤⑦ 全部中断。
#
#  对策：把 SciPy/HiGHS 作为首选引擎 —— 它随 wheel 一起分发、进程内求解、
#  不依赖任何外部 exe，也不依赖 MSVC 运行库；CBC 只在"确实可用"时才用于
#  交叉验证与 MILP。三层降级：PuLP/CBC → SciPy/HiGHS → 纯 Python 规则策略，
#  任何一层不可用都自动切换并记入日志，绝不中断计算。
# ==============================================================================
SOLVER: Dict[str, Any] = {"scipy": False, "cbc": False, "probed": False,
                          "scipy_msg": "", "cbc_msg": "", "rule_fallback": False}


def probe_solvers(force: bool = False) -> Dict[str, Any]:
    """实测各求解器是否真的可用；CBC 需真正启动一次子进程，故用微型 LP 试跑。结果缓存。"""
    if SOLVER["probed"] and not force:
        return SOLVER
    SOLVER["rule_fallback"] = False
    # ---- SciPy / HiGHS：进程内求解，试跑一个最小 LP ----
    if HAS_SCIPY:
        try:
            rr = _linprog(c=[1.0], bounds=[(0.0, 1.0)], method="highs")
            SOLVER["scipy"] = bool(rr.success)
            SOLVER["scipy_msg"] = "" if rr.success else f"HiGHS 返回 {getattr(rr, 'message', '')}"
        except Exception as e:
            SOLVER["scipy"] = False
            SOLVER["scipy_msg"] = f"{type(e).__name__}: {e}"
    else:
        SOLVER["scipy"] = False
        SOLVER["scipy_msg"] = "未安装 scipy"
    # ---- PuLP / CBC：必须真正启动子进程才算可用 ----
    if HAS_PULP:
        try:
            pr = pl.LpProblem("probe", pl.LpMinimize)
            pv = pl.LpVariable("x", 0.0, 1.0)
            pr += pv >= 0.5
            pr += pv
            pr.solve(pl.PULP_CBC_CMD(msg=0, timeLimit=10))
            ok = pl.LpStatus.get(pr.status, "?") == "Optimal"
            SOLVER["cbc"] = ok
            SOLVER["cbc_msg"] = "" if ok else f"CBC 返回 {pl.LpStatus.get(pr.status, '?')}"
        except Exception as e:
            SOLVER["cbc"] = False
            SOLVER["cbc_msg"] = f"{type(e).__name__}: {str(e)[:180]}"
    else:
        SOLVER["cbc"] = False
        SOLVER["cbc_msg"] = "未安装 PuLP"
    SOLVER["probed"] = True
    return SOLVER


def primary_engine() -> str:
    """当前实际可用的首选求解引擎（供界面与日志展示）。"""
    probe_solvers()
    if SOLVER["scipy"]:
        return "SciPy/HiGHS（进程内）"
    if SOLVER["cbc"]:
        return "PuLP/CBC"
    return "纯 Python 规则策略（近似）"


def engine_report() -> List[str]:
    """求解器自检报告。"""
    probe_solvers()
    lines = [
        "求解器自检：SciPy/HiGHS " + ("可用" if SOLVER["scipy"] else
                                 f"不可用（{SOLVER['scipy_msg'][:100]}）"),
        "            PuLP/CBC   " + ("可用" if SOLVER["cbc"] else
                                 f"不可用（{SOLVER['cbc_msg'][:150]}）"),
        f"            → 主求解引擎：{primary_engine()}",
    ]
    if not SOLVER["scipy"]:
        lines.append("建议：pip install scipy —— 进程内求解器，随 wheel 分发、不需要外部 exe，"
                     "最可靠；缺失时网格搜索会明显变慢。")
    if not SOLVER["cbc"]:
        lines.append("说明：CBC 不可用不影响出结果（已自动改用进程内 HiGHS）。"
                     "仅「MILP 硬约束」与「CBC 复算交叉验证」会被跳过。"
                     "如需恢复：安装 Microsoft Visual C++ 2015-2022 (x64) 运行库，"
                     "并把 %TEMP% 加入杀毒软件白名单。")
    return lines


# ==============================================================================
#  §1  工具函数
# ==============================================================================
def fnum(x: Any, nd: int = 4) -> str:
    """安全数值格式化（None / NaN / Inf 统一显示为 '-'）。"""
    try:
        v = float(x)
    except (TypeError, ValueError):
        return "-"
    if math.isnan(v) or math.isinf(v):
        return "-"
    return f"{v:.{nd}f}"


def safe_div(a: float, b: float, default: float = 0.0) -> float:
    try:
        if abs(float(b)) < 1e-12:
            return default
        return float(a) / float(b)
    except (TypeError, ValueError, ZeroDivisionError):
        return default


def clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


ICON_NAME = "Calc.ico"


def icon_path() -> str:
    """定位程序图标：打包(onefile)后从 _MEIPASS 取，源码运行时从脚本目录取。"""
    bases = [getattr(sys, "_MEIPASS", ""), _osp.dirname(_osp.abspath(__file__)), os.getcwd()]
    for b in bases:
        if not b:
            continue
        p = _osp.join(b, ICON_NAME)
        if _osp.exists(p):
            return p
    return ""


# ==============================================================================
#  §2  参数容器
# ==============================================================================
@dataclass
class Spec:
    """全部题目参数与可调参数（默认值 = 题目 3.1 节给定值）。功率 kW，电量 kWh，电价 元/kWh。"""

    # 装机 kW 与最大负荷 kW
    wind_cap: Dict[str, float] = field(default_factory=lambda: {"A": 0.0, "B": 1000.0, "C": 500.0})
    pv_cap: Dict[str, float] = field(default_factory=lambda: {"A": 750.0, "B": 0.0, "C": 600.0})
    max_load: Dict[str, float] = field(default_factory=lambda: {"A": 447.0, "B": 419.0, "C": 506.0})

    # 价格 元/kWh
    price_buy: float = 1.0
    price_wind: float = 0.5
    price_pv: float = 0.4
    price_sell: float = 0.0          # 0 = 富余绿电不允许上网（题目设定）

    # 分时电价
    tou_enable: bool = False
    tou_peak: float = 1.2
    tou_flat: float = 1.0
    tou_valley: float = 0.6

    # 储能（磷酸铁锂）
    cost_p: float = 800.0
    cost_e: float = 1800.0
    eta_c: float = 0.95
    eta_d: float = 0.95
    soc_min: float = 0.10
    soc_max: float = 0.90
    life_y: int = 10

    dt: float = 1.0
    days: int = 365

    def buy_series(self, hours: np.ndarray) -> np.ndarray:
        """逐时购电电价数组（分时电价时按峰/平/谷划分）。"""
        if not self.tou_enable:
            return np.full(len(hours), float(self.price_buy))
        out = np.empty(len(hours), dtype=float)
        for i, h in enumerate(hours):
            h = int(h)
            if h in HOURS_PEAK:
                out[i] = self.tou_peak
            elif h in HOURS_VALLEY:
                out[i] = self.tou_valley
            else:
                out[i] = self.tou_flat
        return out

    def clone(self, **kw) -> "Spec":
        d = {k: (dict(v) if isinstance(v, dict) else v) for k, v in self.__dict__.items()}
        d.update(kw)
        return Spec(**d)


# ==============================================================================
#  §3  数据层
# ==============================================================================
# ==============================================================================
#  §3.0  内置数据模板（默认数据源）
#  ----------------------------------------------------------------------------
#  下列数值由题目三个附件导出并固化在代码内，使本程序在没有任何外部文件时
#  也能独立运行（成为真正的单文件程序）。
#    load : 附件1 三园区典型日逐时负荷 (kW)，24 点
#    pu2  : 附件2 典型日风光出力 (p.u.)，以该园区该电源额定装机为基准
#    m_pu : 附件3 12 个月典型日风光出力 (p.u.)，每行 24 点
#  界面可「载入 Excel」覆盖本模板，也可用「数据调整」按比例缩放。
# ==============================================================================
_EMBEDDED = {
    "load": {
        "A": [
            275.0, 275.0, 277.0, 310.0, 310.0, 293.0, 293.0, 380.0,
            375.0, 281.0, 447.0, 447.0, 447.0, 405.0, 404.0, 403.0,
            268.0, 313.0, 287.0, 288.0, 284.0, 287.0, 277.0, 275.0,
        ],
        "B": [
            241.0, 253.0, 329.0, 315.0, 290.0, 270.0, 307.0, 354.0,
            264.0, 315.0, 313.0, 291.0, 360.0, 369.0, 389.0, 419.0,
            412.0, 291.0, 379.0, 303.0, 331.0, 306.0, 285.0, 324.0,
        ],
        "C": [
            302.0, 292.0, 307.0, 293.0, 271.0, 252.0, 283.0, 223.0,
            292.0, 283.0, 287.0, 362.0, 446.0, 504.0, 455.0, 506.0,
            283.0, 311.0, 418.0, 223.0, 229.0, 361.0, 302.0, 291.0,
        ],
    },
    "pu2": {
        "A_pv": [
            0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0058,
            0.3026, 0.6020, 0.7711, 0.8555, 0.8531, 0.7842, 0.6437, 0.4242,
            0.0619, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
        ],
        "B_wd": [
            0.2301, 0.3828, 0.2968, 0.4444, 0.5029, 0.3609, 0.2402, 0.0473,
            0.1538, 0.1068, 0.0518, 0.2169, 0.3546, 0.2194, 0.1110, 0.2186,
            0.3779, 0.3421, 0.5008, 0.4646, 0.2197, 0.1783, 0.1535, 0.0000,
        ],
        "C_pv": [
            0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0105,
            0.3280, 0.6314, 0.7936, 0.8925, 0.8999, 0.8221, 0.6667, 0.4275,
            0.0216, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
        ],
        "C_wd": [
            0.1464, 0.2175, 0.3959, 0.1831, 0.4716, 0.6215, 0.2946, 0.1214,
            0.0250, 0.3023, 0.0196, 0.1224, 0.3335, 0.2653, 0.1220, 0.1633,
            0.2645, 0.3408, 0.3183, 0.3299, 0.1703, 0.1655, 0.1897, 0.2323,
        ],
    },
    "m_pu": {
        "A_pv": [
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0211, 0.1386, 0.1292, 0.2718,
                0.2833, 0.1637, 0.1588, 0.0388, 0.0056, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0058, 0.3026, 0.6020, 0.7711, 0.8555,
                0.8531, 0.7842, 0.6437, 0.4242, 0.0619, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0019, 0.0441, 0.0566, 0.4737, 0.7593, 0.3318,
                0.3298, 0.2556, 0.5223, 0.5206, 0.1773, 0.0286, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0512, 0.3066, 0.5595, 0.5139, 0.2840, 0.2183,
                0.2107, 0.2017, 0.1746, 0.1491, 0.1903, 0.0683, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0211, 0.1325, 0.3657, 0.5905, 0.6590, 0.3411, 0.3372,
                0.3511, 0.3952, 0.3482, 0.5067, 0.2182, 0.1039, 0.0164, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0505, 0.0805, 0.1134, 0.0674, 0.0524, 0.0665,
                0.0888, 0.0620, 0.0677, 0.0295, 0.0064, 0.0052, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0497, 0.0766, 0.2090, 0.3335, 0.5358, 0.7172, 0.2879,
                0.2684, 0.2838, 0.1983, 0.1171, 0.1009, 0.0626, 0.0131, 0.0035, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0211, 0.0872, 0.2451, 0.4992, 0.5054, 0.4558, 0.4762,
                0.5406, 0.6594, 0.4566, 0.4732, 0.2630, 0.1035, 0.0211, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0030, 0.0983, 0.3506, 0.5923, 0.7523, 0.8269, 0.4139,
                0.2192, 0.1928, 0.2923, 0.4767, 0.2566, 0.0649, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0628, 0.3095, 0.4607, 0.2066, 0.1725, 0.1647,
                0.4784, 0.4979, 0.4695, 0.4631, 0.1981, 0.0168, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0032, 0.2198, 0.5053, 0.6994, 0.8103, 0.8442,
                0.8112, 0.7111, 0.5348, 0.2839, 0.0290, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0502, 0.2638, 0.1654, 0.1472,
                0.2116, 0.1384, 0.0851, 0.0346, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
        ],
        "B_wd": [
            [
                0.0383, 0.0252, 0.0428, 0.1136, 0.1510, 0.1858, 0.1681, 0.1374, 0.1321, 0.4063, 0.5659, 0.4769,
                0.5536, 0.3979, 0.3825, 0.5877, 0.3733, 0.0560, 0.0727, 0.0000, 0.6115, 0.3745, 0.2078, 0.1141,
            ],
            [
                0.2301, 0.3828, 0.2968, 0.4444, 0.5029, 0.3609, 0.2402, 0.0473, 0.1538, 0.1068, 0.0518, 0.2169,
                0.3546, 0.2194, 0.1110, 0.2186, 0.3779, 0.3421, 0.5008, 0.4646, 0.2197, 0.1783, 0.1535, 0.4347,
            ],
            [
                0.2489, 0.2541, 0.2369, 0.3412, 0.2975, 0.2502, 0.1299, 0.0921, 0.0034, 0.0000, 0.0470, 0.2288,
                0.3012, 0.4844, 0.4435, 0.2366, 0.2947, 0.4051, 0.5369, 0.7929, 0.4552, 0.5241, 0.7451, 0.7863,
            ],
            [
                0.4867, 0.4853, 0.7585, 0.8324, 0.7830, 0.8216, 0.4912, 0.0637, 0.0000, 0.1329, 0.3295, 0.3339,
                0.4924, 0.4271, 0.4166, 0.1031, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0240,
            ],
            [
                0.2230, 0.1507, 0.1773, 0.1710, 0.1731, 0.1427, 0.1353, 0.0797, 0.0556, 0.0134, 0.0000, 0.0051,
                0.0077, 0.0799, 1.0000, 0.9503, 0.5921, 0.3566, 0.5398, 0.5925, 0.3965, 0.2686, 0.6659, 0.4621,
            ],
            [
                0.4311, 0.3796, 0.3780, 0.3851, 0.4490, 0.4555, 0.5864, 0.6373, 0.6523, 0.4469, 0.3470, 0.4006,
                0.2806, 0.2811, 0.2485, 0.5020, 0.4498, 0.0419, 0.0000, 0.0000, 0.2218, 0.3322, 0.2264, 0.1646,
            ],
            [
                0.9732, 0.6883, 0.0000, 0.4278, 0.3182, 0.0731, 0.3966, 0.6820, 0.5707, 0.3023, 0.1431, 0.0564,
                0.0064, 0.0000, 0.0037, 0.0002, 0.0060, 0.0000, 0.0000, 0.0001, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0308, 0.0503, 0.1114, 0.0459, 0.2058, 0.2724, 0.4302, 0.8621, 0.7118, 0.6684,
                0.4070, 0.3716, 0.3446, 0.2205, 0.0754, 0.0102, 0.0321, 0.0473, 0.0759, 0.0389, 0.0000, 0.0000,
            ],
            [
                0.3043, 0.3569, 0.3506, 0.2591, 0.2508, 0.1438, 0.1448, 0.0477, 0.0272, 0.0385, 0.0066, 0.0000,
                0.0000, 0.0000, 0.0000, 0.0587, 0.2062, 0.2513, 0.2046, 0.2000, 0.1987, 0.0791, 0.0261, 0.1216,
            ],
            [
                0.0022, 0.0314, 0.0396, 0.0262, 0.0175, 0.0206, 0.0000, 0.0049, 0.0000, 0.5745, 0.1809, 0.0788,
                0.1182, 0.5649, 0.4118, 0.2934, 0.5798, 0.0308, 0.0121, 0.0933, 0.1388, 0.0651, 0.0161, 0.1638,
            ],
            [
                0.0672, 0.0480, 0.0576, 0.0876, 0.0236, 0.0565, 0.0844, 0.0174, 0.0038, 0.0043, 0.0000, 0.0170,
                0.1171, 0.0981, 0.1885, 0.3279, 0.3303, 0.2437, 0.1299, 0.1429, 0.0571, 0.0007, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0103, 0.0656, 0.0610, 0.0273, 0.0054, 0.0080, 0.0302, 0.0067, 0.0301, 0.1545,
                0.1629, 0.0674, 0.1250, 0.0858, 0.1916, 0.3483, 0.5590, 0.5902, 0.4001, 0.2011, 0.0365, 0.0040,
            ],
        ],
        "C_wd": [
            [
                0.0776, 0.0227, 0.0288, 0.0892, 0.1176, 0.1653, 0.1858, 0.1482, 0.0857, 0.3595, 0.4915, 0.4305,
                0.5039, 0.7286, 0.7091, 0.6739, 0.6514, 0.1197, 0.0000, 0.4937, 0.3359, 0.4553, 0.3725, 0.1493,
            ],
            [
                0.1464, 0.2175, 0.3959, 0.1831, 0.4716, 0.6215, 0.2946, 0.1214, 0.0250, 0.3023, 0.0196, 0.1224,
                0.3335, 0.2653, 0.1220, 0.1633, 0.2645, 0.3408, 0.3183, 0.3299, 0.1703, 0.1655, 0.1897, 0.2323,
            ],
            [
                0.2435, 0.1934, 0.2144, 0.2790, 0.3603, 0.2336, 0.1462, 0.1120, 0.0228, 0.0000, 0.0000, 0.0964,
                0.2539, 0.3507, 0.4616, 0.3220, 0.1931, 0.3921, 0.4637, 0.3234, 0.3862, 0.6747, 0.6660, 0.8572,
            ],
            [
                0.4624, 0.4531, 0.5662, 0.7941, 0.6864, 0.8022, 0.6526, 0.2171, 0.0000, 0.0980, 0.2403, 0.2924,
                0.4047, 0.4297, 0.4150, 0.1393, 0.0622, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.3490, 0.1805, 0.1831, 0.2046, 0.2111, 0.1357, 0.1586, 0.0885, 0.0144, 0.0287, 0.0000, 0.0049,
                0.0418, 0.0024, 0.5427, 1.0000, 0.7214, 0.3210, 0.3978, 0.4790, 0.3863, 0.2731, 0.3603, 0.4504,
            ],
            [
                0.3174, 0.3853, 0.3420, 0.3707, 0.4206, 0.3412, 0.4821, 0.5134, 0.5851, 0.4582, 0.2762, 0.3051,
                0.2358, 0.2523, 0.1986, 0.3112, 0.4974, 0.2380, 0.0000, 0.0000, 0.0000, 0.3450, 0.2220, 0.1727,
            ],
            [
                0.9441, 0.7945, 0.2934, 0.0200, 0.7095, 0.0126, 0.2527, 0.3751, 0.5673, 0.3329, 0.1229, 0.0162,
                0.0010, 0.0000, 0.0000, 0.0008, 0.0045, 0.0058, 0.0000, 0.0008, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0018, 0.0484, 0.0575, 0.0986, 0.1156, 0.2413, 0.2050, 0.6214, 0.8263, 0.6222,
                0.4375, 0.2540, 0.3665, 0.1679, 0.0917, 0.0207, 0.0224, 0.0378, 0.0448, 0.0244, 0.0070, 0.0000,
            ],
            [
                0.2098, 0.2675, 0.2676, 0.2855, 0.2310, 0.1568, 0.1524, 0.0982, 0.0563, 0.0139, 0.0040, 0.0000,
                0.0000, 0.0000, 0.0000, 0.0520, 0.2422, 0.3603, 0.3717, 0.2859, 0.2971, 0.1909, 0.0899, 0.0434,
            ],
            [
                0.0004, 0.0031, 0.0436, 0.0331, 0.0198, 0.0303, 0.0025, 0.0336, 0.0003, 0.3182, 0.3729, 0.0327,
                0.0877, 0.4140, 0.6525, 0.2853, 0.5620, 0.1277, 0.0267, 0.0753, 0.1260, 0.1383, 0.0559, 0.1581,
            ],
            [
                0.0938, 0.0555, 0.0751, 0.0628, 0.0466, 0.1320, 0.0414, 0.0258, 0.0000, 0.0378, 0.0000, 0.0000,
                0.0833, 0.1109, 0.1436, 0.4564, 0.4984, 0.3027, 0.1701, 0.1534, 0.1717, 0.0564, 0.0001, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0015, 0.0037, 0.0622, 0.0468, 0.0063, 0.0169, 0.0108, 0.0081, 0.0106, 0.0574,
                0.1868, 0.0636, 0.0878, 0.1416, 0.0630, 0.1778, 0.3990, 0.4948, 0.3475, 0.1897, 0.0383, 0.0099,
            ],
        ],
        "C_pv": [
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0006, 0.0452, 0.1109, 0.2398, 0.2260,
                0.2097, 0.1661, 0.0657, 0.0156, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0105, 0.3280, 0.6314, 0.7936, 0.8925,
                0.8999, 0.8221, 0.6667, 0.4275, 0.0216, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0458, 0.0793, 0.2139, 0.8152, 0.5905,
                0.1979, 0.1533, 0.4327, 0.4309, 0.2247, 0.0053, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0017, 0.0625, 0.3230, 0.5827, 0.7123, 0.3472, 0.2666,
                0.2651, 0.1256, 0.1046, 0.1320, 0.1870, 0.0645, 0.0012, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0306, 0.1499, 0.2876, 0.4056, 0.7822, 0.8298, 0.6639,
                0.6435, 0.8287, 0.4290, 0.2985, 0.1902, 0.0535, 0.0145, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0002, 0.0549, 0.0516, 0.0835, 0.3040, 0.2947, 0.2759, 0.1660,
                0.0751, 0.0610, 0.1417, 0.1453, 0.1080, 0.0409, 0.0076, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0043, 0.0339, 0.1233, 0.3373, 0.4640, 0.5592, 0.2561, 0.1372,
                0.1730, 0.1361, 0.1862, 0.2402, 0.2191, 0.1281, 0.0572, 0.0066, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0007, 0.0932, 0.1107, 0.2254, 0.3650, 0.5285, 0.2516,
                0.3185, 0.6565, 0.4312, 0.4048, 0.1940, 0.0838, 0.0143, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0032, 0.0920, 0.3139, 0.5144, 0.6583, 0.7295, 0.4107,
                0.1401, 0.2290, 0.2378, 0.4258, 0.2167, 0.0361, 0.0002, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0645, 0.2847, 0.3920, 0.1967, 0.1974, 0.1494,
                0.4760, 0.4606, 0.5619, 0.3966, 0.1704, 0.0124, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0030, 0.2272, 0.4818, 0.6489, 0.7459, 0.7774,
                0.7507, 0.6622, 0.5125, 0.2924, 0.0173, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
            [
                0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0036, 0.2421, 0.6494, 0.7908, 0.8474,
                0.8253, 0.6944, 0.5535, 0.0356, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000,
            ],
        ],
    },
    "caps": {"wind": {"A": 0.0, "B": 1000.0, "C": 500.0},
             "pv": {"A": 750.0, "B": 0.0, "C": 600.0}},
    "hours": [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23],
}


HERE = _osp.dirname(_osp.abspath(__file__))
CAND_DIRS = [HERE, _osp.join(HERE, "microgrid_B"), os.getcwd()]

F_LOAD = ["附件1：各园区典型日负荷数据.xlsx", "附件1：各园区典型日负荷数据.xlxs"]
F_GEN2 = ["att2.xlsx", "附件2：各园区典型日风光发电数据.xlsx"]
F_GEN3 = ["附件3：12个月各园区典型日风光发电数据_原.xlsx",
          "附件3：12个月各园区典型日风光发电数据.xlsx"]


def _find(names: Sequence[str]) -> Optional[str]:
    for d in CAND_DIRS:
        for n in names:
            p = _osp.join(d, n)
            if _osp.exists(p):
                return p
    return None


def _sheet_rows(path: str) -> List[list]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = [["" if c is None else c for c in r]
            for r in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True)]
    wb.close()
    return rows


def _as_hour(v) -> int:
    if hasattr(v, "hour"):
        return int(v.hour)
    if isinstance(v, str):
        return int(str(v).split(":")[0])
    return int(v)


def _num(v, default=0.0) -> float:
    if isinstance(v, (int, float)):
        return float(v)
    if v == "" or v is None:
        return default
    try:
        return float(v)
    except Exception:
        return default


# 每个园区的电源构成：风电 / 光伏各自对应的标幺序列键（None = 该园区无此电源）
PU_SRC = {"A": {"wind": None, "pv": "A_pv"},
          "B": {"wind": "B_wd", "pv": None},
          "C": {"wind": "C_wd", "pv": "C_pv"}}


@dataclass
class Data:
    """数据容器。功率 kW、电量 kWh（Δt = 1 h）。
    标幺序列与装机容量分开保存，因此界面调整装机后无需重读文件即可重算。"""
    hours: np.ndarray
    load: Dict[str, np.ndarray]                 # 三园区典型日逐时负荷 (kW)
    pu2: Dict[str, np.ndarray]                  # 典型日风光出力 (p.u.)
    m_pu: Dict[str, np.ndarray]                 # 12 个月典型日风光出力 (p.u.)，(12,24)
    caps: Dict[str, Dict[str, float]]           # {"wind": {...}, "pv": {...}} 当前装机 (kW)
    source: str = "内置数据模板"
    checks: List[str] = field(default_factory=list)
    src_checks: List[Dict[str, Any]] = field(default_factory=list)

    def _kws(self, park: str, month: int, kind: str) -> np.ndarray:
        """标幺值 × 当前装机 → kW。"""
        key = PU_SRC[park][kind]
        if key is None:
            return np.zeros(24)
        arr = self.pu2[key] if month <= 0 else self.m_pu[key][month - 1]
        return np.asarray(arr, float) * float(self.caps[kind][park])

    def arrays(self, park: str, month: int = 0) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
        """返回 (负荷, 风电, 光伏)，单位 kW。month=0 → 附件2 典型日；1..12 → 附件3 对应月份。"""
        return (np.asarray(self.load[park], float),
                self._kws(park, month, "wind"), self._kws(park, month, "pv"))

    def joint_arrays(self, month: int = 0) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
        L = sum(np.asarray(self.load[p], float) for p in PARKS)
        W = sum(self._kws(p, month, "wind") for p in PARKS)
        V = sum(self._kws(p, month, "pv") for p in PARKS)
        return L, W, V

    @property
    def gen(self) -> Dict[str, np.ndarray]:
        return {p: self._kws(p, 0, "wind") + self._kws(p, 0, "pv") for p in PARKS}

    @property
    def net(self) -> Dict[str, np.ndarray]:
        return {p: np.asarray(self.load[p], float) - self.gen[p] for p in PARKS}


def _assemble(load: Dict[str, np.ndarray], pu2: Dict[str, np.ndarray],
              m_pu: Dict[str, np.ndarray], hours: np.ndarray, hours2: np.ndarray,
              caps: Dict[str, Dict[str, float]], source: str, n2: int, n3: int) -> Data:
    """由已解析的负荷/标幺数据装配 Data，并生成数据基本检查项。"""
    sp = Spec()
    data = Data(hours=hours, load=load, pu2=pu2, m_pu=m_pu, caps=caps, source=source)
    wind_kw = {p: data._kws(p, 0, "wind") for p in PARKS}
    pv_kw = {p: data._kws(p, 0, "pv") for p in PARKS}

    rows: List[Dict[str, Any]] = []
    checks: List[str] = []

    def add(item: str, value: str) -> None:
        rows.append({"检查项": item, "结果": value})
        checks.append(f"【{item}】{value}")

    add("数据来源", source)
    add("附件1 数据行数 / 粒度",
        f"{len(hours)} 行，时间范围 {hours.min():02d}:00–{hours.max():02d}:00，粒度 1 h")
    n_nan = int(sum(int(np.isnan(load[p]).sum()) for p in PARKS))
    add("附件1 缺失值", f"{n_nan} 个")
    for p in PARKS:
        ok = abs(load[p].max() - sp.max_load[p]) < 0.5
        add(f"附件1 园区{p} 峰值/日电量/负荷率",
            f"峰值 {load[p].max():.0f} kW（题目 {sp.max_load[p]:.0f}，{'一致' if ok else '不一致'}）"
            f"，日电量 {load[p].sum():.1f} kWh，负荷率 {100 * load[p].mean() / load[p].max():.1f}%")
    add("附件2 数据行数", f"{n2} 行，时间轴与附件1 "
                          f"{'一致' if np.array_equal(hours, hours2) else '不一致'}")
    for k, v in pu2.items():
        add(f"附件2 {k} 标幺范围",
            f"min={v.min():.4f} max={v.max():.4f} 均值={v.mean():.4f}，"
            f"越界点 {int(((v < -1e-9) | (v > 1 + 1e-9)).sum())} 个")
    add("附件3 数据量", f"12 月 × 24 时点 × 4 序列 = {n3} 条记录")
    # 附件2 与附件3 逐序列最优匹配
    for name, a, key in (("园区A光伏", pu2["A_pv"], "A_pv"), ("园区B风电", pu2["B_wd"], "B_wd"),
                         ("园区C风电", pu2["C_wd"], "C_wd"), ("园区C光伏", pu2["C_pv"], "C_pv")):
        cnt = np.array([int((np.abs(a - m_pu[key][m]) > 1e-9).sum()) for m in range(12)])
        dmax = np.array([np.abs(a - m_pu[key][m]).max() for m in range(12)])
        best = int(np.lexsort((dmax, cnt))[0])
        dvec = np.abs(a - m_pu[key][best])
        bad = [int(i) for i in np.nonzero(dvec > 1e-9)[0]]
        tail = "完全一致" if not bad else f"差异 t={bad}，最大 {dmax[best]:.4f} p.u."
        add(f"一致性 附件2 与 附件3 {name}",
            f"最吻合 {best + 1} 月（差异 {cnt[best]}/24 点），{tail}")
    night = list(range(0, 6)) + [22, 23]
    add("物理合理性 光伏夜间出力", f"园区A 光伏夜间全为 0：{bool(np.allclose(pv_kw['A'][night], 0))}")
    add("物理合理性 装机对应", f"园区B无光伏、园区A无风电："
                              f"{bool(np.allclose(pv_kw['B'], 0) and np.allclose(wind_kw['A'], 0))}")
    add("当前装机 (kW)", "风电 " + " / ".join(f"{p}:{caps['wind'][p]:.0f}" for p in PARKS)
        + "；光伏 " + " / ".join(f"{p}:{caps['pv'][p]:.0f}" for p in PARKS))

    data.checks, data.src_checks = checks, rows
    return data


def embedded_data() -> Data:
    """由代码内置的数据模板装配（默认数据源，不读取任何外部文件）。"""
    e = _EMBEDDED
    load = {p: np.array(e["load"][p], float) for p in PARKS}
    pu2 = {k: np.array(v, float) for k, v in e["pu2"].items()}
    m_pu = {k: np.array(v, float) for k, v in e["m_pu"].items()}
    hours = np.array(e["hours"], dtype=int)
    caps = {"wind": dict(e["caps"]["wind"]), "pv": dict(e["caps"]["pv"])}
    return _assemble(load, pu2, m_pu, hours, hours.copy(), caps,
                     "内置数据模板（代码内置，无需外部文件）", len(hours), 12 * 24 * 4)


def excel_data(paths: Optional[Sequence[str]] = None) -> Data:
    """
    从 Excel 读取数据。
    paths=None 时自动在 [脚本目录, 脚本目录/microgrid_B, 当前目录] 中查找三个附件；
    也可显式给定 [附件1, 附件2, 附件3] 三个路径（缺省的后两个将回退为内置模板对应部分）。
    """
    if not HAS_XL:
        raise RuntimeError("缺少 openpyxl，无法读取 Excel（pip install openpyxl）")
    if paths:
        ps = list(paths)
        p1 = ps[0]
        p2 = ps[1] if len(ps) > 1 else None
        p3 = ps[2] if len(ps) > 2 else None
    else:
        p1, p2, p3 = _find(F_LOAD), _find(F_GEN2), _find(F_GEN3)
    if p1 is None:
        raise FileNotFoundError(f"未找到附件1（已搜索 {CAND_DIRS}）")
    if p2 is None:
        p2 = _find(F_GEN2)
    if p3 is None:
        p3 = _find(F_GEN3)
    if p2 is None or p3 is None:
        # 缺附件2/3 时以内置模板补齐，只替换用户真正提供的部分
        emb = embedded_data()
        pu2 = emb.pu2
        m_pu = emb.m_pu
        hours2 = emb.hours
        n2, n3 = len(emb.hours), 12 * 24 * 4
        if p2 is not None:
            r2 = _sheet_rows(p2)
            h2 = next(i for i, r in enumerate(r2) if r and str(r[0]).startswith("时间"))
            b2 = [r for r in r2[h2 + 1:] if r and r[0] != ""]
            pu2 = {"A_pv": np.array([_num(r[1]) for r in b2], float),
                   "B_wd": np.array([_num(r[2]) for r in b2], float),
                   "C_pv": np.array([_num(r[3]) for r in b2], float),
                   "C_wd": np.array([_num(r[4]) for r in b2], float)}
            hours2 = np.array([_as_hour(r[0]) for r in b2], dtype=int)
            n2 = len(b2)
        if p3 is not None:
            m_pu = _read_months(p3)
        src = "Excel（部分附件缺失，已用内置模板补齐）"
        sp = Spec()
        caps = {"wind": dict(sp.wind_cap), "pv": dict(sp.pv_cap)}
        return _excel_load(p1, pu2, m_pu, hours2, caps, src, n2, n3)

    pu2, hours2, n2 = _read_gen2(p2)
    m_pu = _read_months(p3)
    sp = Spec()
    caps = {"wind": dict(sp.wind_cap), "pv": dict(sp.pv_cap)}
    src = f"Excel：{_osp.basename(p1)} 等 3 个附件"
    return _excel_load(p1, pu2, m_pu, hours2, caps, src, n2, 12 * 24 * 4)


def _read_gen2(p2: str):
    r2 = _sheet_rows(p2)
    h2 = next(i for i, r in enumerate(r2) if r and str(r[0]).startswith("时间"))
    b2 = [r for r in r2[h2 + 1:] if r and r[0] != ""]
    pu = {"A_pv": np.array([_num(r[1]) for r in b2], float),
          "B_wd": np.array([_num(r[2]) for r in b2], float),
          "C_pv": np.array([_num(r[3]) for r in b2], float),
          "C_wd": np.array([_num(r[4]) for r in b2], float)}
    return pu, np.array([_as_hour(r[0]) for r in b2], dtype=int), len(b2)


def _read_months(p3: str) -> Dict[str, np.ndarray]:
    r3 = _sheet_rows(p3)
    h3 = next(i for i, r in enumerate(r3) if r and str(r[0]).startswith("时间"))
    b3 = [r for r in r3[h3 + 1:] if r and r[0] != ""][:24]
    m = {"A_pv": np.zeros((12, 24)), "B_wd": np.zeros((12, 24)),
         "C_wd": np.zeros((12, 24)), "C_pv": np.zeros((12, 24))}
    for mi in range(12):
        c0 = 1 + 4 * mi                     # 附件3 列序：A_pv, B_wd, C_wd, C_pv
        for t in range(24):
            m["A_pv"][mi, t] = _num(b3[t][c0 + 0])
            m["B_wd"][mi, t] = _num(b3[t][c0 + 1])
            m["C_wd"][mi, t] = _num(b3[t][c0 + 2])
            m["C_pv"][mi, t] = _num(b3[t][c0 + 3])
    return m


def _excel_load(p1, pu2, m_pu, hours2, caps, src, n2, n3) -> Data:
    r1 = _sheet_rows(p1)
    h1 = next(i for i, r in enumerate(r1) if r and str(r[0]).startswith("时间"))
    b1 = [r for r in r1[h1 + 1:] if r and r[0] != ""]
    hours = np.array([_as_hour(r[0]) for r in b1], dtype=int)
    load = {"A": np.array([_num(r[1]) for r in b1], float),
            "B": np.array([_num(r[2]) for r in b1], float),
            "C": np.array([_num(r[3]) for r in b1], float)}
    return _assemble(load, pu2, m_pu, hours, hours2, caps, src, n2, n3)


def load_data(source: str = "embedded", paths: Optional[Sequence[str]] = None,
              verbose: bool = False) -> Data:
    """统一数据入口。source='embedded'（默认：代码内置模板）或 'excel'（读取附件文件）。"""
    d = embedded_data() if source != "excel" else excel_data(paths)
    if verbose:
        for c in d.checks:
            print(c)
    return d


# ==============================================================================
#  §4  优化模型层
# ==============================================================================
@dataclass
class OpResult:
    """一次调度（基线 / 规则 / LP / 联合）的完整结果。"""
    park: str = ""
    label: str = ""
    P: float = 0.0
    E: float = 0.0
    engine: str = ""
    status: str = ""
    wu: np.ndarray = field(default_factory=lambda: np.zeros(24))   # 风电入系统
    vu: np.ndarray = field(default_factory=lambda: np.zeros(24))   # 光伏入系统
    ws: np.ndarray = field(default_factory=lambda: np.zeros(24))   # 风电外送
    ps: np.ndarray = field(default_factory=lambda: np.zeros(24))   # 光伏外送
    buy: np.ndarray = field(default_factory=lambda: np.zeros(24))
    ch: np.ndarray = field(default_factory=lambda: np.zeros(24))
    dis: np.ndarray = field(default_factory=lambda: np.zeros(24))
    soc: np.ndarray = field(default_factory=lambda: np.zeros(24))
    e0: float = 0.0

    def metrics(self, L, W, V, sp: Spec) -> Dict[str, float]:
        L = np.asarray(L, float); W = np.asarray(W, float); V = np.asarray(V, float)
        dt = sp.dt
        e_load = float(L.sum() * dt)
        e_gen = float((W + V).sum() * dt)
        e_wu = float(self.wu.sum() * dt)
        e_vu = float(self.vu.sum() * dt)
        e_ws = float(self.ws.sum() * dt)
        e_ps = float(self.ps.sum() * dt)
        e_green = e_wu + e_vu
        e_sell = e_ws + e_ps
        e_curt = e_gen - e_green - e_sell
        e_buy = float(self.buy.sum() * dt)
        e_ch = float(self.ch.sum() * dt)
        e_dis = float(self.dis.sum() * dt)
        c_green = (e_wu + e_ws) * sp.price_wind + (e_vu + e_ps) * sp.price_pv
        c_buy = float(np.dot(np.asarray(self.buy, float), sp.buy_series(np.arange(len(self.buy)))) * dt)
        c_sell = e_sell * sp.price_sell
        c_tot = c_green + c_buy - c_sell
        return {
            "负荷电量": e_load, "新能源发电量": e_gen, "绿电自用电量": e_green,
            "其中风电自用": e_wu, "其中光伏自用": e_vu, "购电量": e_buy, "弃电量": e_curt,
            "售电量": e_sell, "充电量": e_ch, "放电量": e_dis,
            "绿电成本": c_green, "购电成本": c_buy, "售电收益": c_sell,
            "总供电成本": c_tot,
            "单位电量平均供电成本": safe_div(c_tot, e_load),
            "新能源渗透率": safe_div(e_green, e_load),
            "弃电率": safe_div(e_curt, e_gen),
            "储能日循环效率": safe_div(e_dis, e_ch),
        }


def _no_storage_dispatch(L, W, V, sp: Spec):
    """无储能确定性调度：绿电优先，光伏(0.4)先于风电(0.5)消纳；盈余按价格决定售/弃。"""
    L = np.asarray(L, float); W = np.asarray(W, float); V = np.asarray(V, float)
    T = len(L)
    surplus = np.maximum(W + V - L, 0.0)
    wu = np.maximum(W - surplus, 0.0)          # 先弃风电（贵）
    rest = np.maximum(surplus - W, 0.0)
    vu = np.maximum(V - rest, 0.0)
    buy = np.maximum(L - W - V, 0.0)
    if sp.price_sell > 0:                       # 允许售电：盈余全部外送
        ws = np.maximum(W - wu, 0.0)
        ps = np.maximum(V - vu, 0.0)
    else:
        ws = np.zeros(T); ps = np.zeros(T)
    return wu, vu, ws, ps, buy


def baseline(L, W, V, sp: Spec, park: str = "") -> OpResult:
    """无储能基线（3.3 节第 4 条，确定性算术）。"""
    r = OpResult(park=park, label="无储能", engine="确定性", status="基线")
    r.wu, r.vu, r.ws, r.ps, r.buy = _no_storage_dispatch(L, W, V, sp)
    return r


def solve_operation(L, W, V, P: float, E: float, sp: Spec, park: str = "",
                    milp: bool = False, time_limit: int = 30) -> OpResult:
    """
    给定储能 (P,E) 的日循环稳态最优运行（调度入口，带自动降级）。
    引擎优先级：PuLP/CBC（确实可用时）→ SciPy/HiGHS（进程内）→ 纯 Python 规则策略。
    任一层不可用或调用失败都会自动切换并记入日志，不向上抛异常。
    """
    L = np.asarray(L, float); W = np.asarray(W, float); V = np.asarray(V, float)
    if P <= 1e-9 or E <= 1e-9:
        r = baseline(L, W, V, sp, park)
        r.label, r.P, r.E = "LP最优(退化)", P, E
        r.status = "无储能退化"
        return r

    probe_solvers()
    err = ""
    if SOLVER["cbc"]:
        try:
            return _pulp_solve(L, W, V, P, E, sp, park=park, milp=milp,
                               time_limit=time_limit)
        except Exception as e:                 # PulpSolverError / 子进程无法启动等
            err = f"{type(e).__name__}: {str(e)[:160]}"
            SOLVER["cbc"] = False
            SOLVER["cbc_msg"] = err

    # ---- 降级 1：进程内 HiGHS（与 PuLP 模型等价，同样是全局最优 LP 解）----
    if SOLVER["scipy"]:
        r = fast_solve(L, W, V, P, E, sp)
        if r is not None:
            r.park, r.label = park, "LP最优"
            r.engine = "LP(HiGHS)"
            r.status = "最优解"
            if err:
                r.status += "｜CBC 不可用，已自动降级为 HiGHS"
            if milp:
                r.status += "｜MILP 不可用，已按 LP 处理"
            return r

    # ---- 降级 2：纯 Python 规则策略（题目设定下即精确最优，详见论文 5.3 节）----
    SOLVER["rule_fallback"] = True
    r = rule_strategy(L, W, V, P, E, sp, park)
    r.label = "规则策略"
    r.engine = "RULE(降级)"
    r.status = "周期稳态（无可用 LP 求解器，结果为规则策略）"
    return r


def _pulp_solve(L, W, V, P: float, E: float, sp: Spec, park: str = "",
                milp: bool = False, time_limit: int = 30) -> OpResult:
    """
    PuLP + CBC 建模求解（CBC 不可用时抛异常，由 solve_operation 降级处理）。
    变量：wu, vu（入系统）、ws, ps（外送）、buy、ch、dis、E、E0。
    约束：功率平衡 + 源出力上限 + SOC 递推 + 日循环稳态 +（可选）充放电互斥。
    目标：min Σ (绿电成本 + 分时购电成本 − 售电收益)。
    """
    L = np.asarray(L, float); W = np.asarray(W, float); V = np.asarray(V, float)
    T = len(L)
    if not HAS_PULP:
        raise RuntimeError("未检测到 PuLP，无法执行 LP 最优调度（pip install pulp）")

    if P <= 1e-9 or E <= 1e-9:
        r = baseline(L, W, V, sp, park)
        r.label, r.P, r.E = "LP最优(退化)", P, E
        r.status = "无储能退化"
        return r

    buy_p = sp.buy_series(np.asarray(range(T)))
    lo, hi = sp.soc_min * E, sp.soc_max * E
    pr = pl.LpProblem(f"op_{park or 'x'}", pl.LpMinimize)
    wu = pl.LpVariable.dicts("wu", range(T), 0, None)
    vu = pl.LpVariable.dicts("vu", range(T), 0, None)
    ws = pl.LpVariable.dicts("ws", range(T), 0, None)
    ps = pl.LpVariable.dicts("ps", range(T), 0, None)
    by = pl.LpVariable.dicts("by", range(T), 0, None)
    ch = pl.LpVariable.dicts("ch", range(T), 0, P)
    dis = pl.LpVariable.dicts("dis", range(T), 0, P)
    Ev = pl.LpVariable.dicts("E", range(T), lo, hi)
    E0 = pl.LpVariable("E0", lo, hi)

    for t in range(T):
        pr += wu[t] <= W[t]
        pr += vu[t] <= V[t]
        pr += wu[t] + ws[t] <= W[t]
        pr += vu[t] + ps[t] <= V[t]
        pr += wu[t] + vu[t] + dis[t] + by[t] == L[t] + ch[t]
        prev = E0 if t == 0 else Ev[t - 1]
        pr += Ev[t] == prev + sp.eta_c * ch[t] * sp.dt - dis[t] * sp.dt / sp.eta_d
    pr += Ev[T - 1] == E0

    if milp:
        z = pl.LpVariable.dicts("z", range(T), 0, 1, cat=pl.LpBinary)
        for t in range(T):
            pr += ch[t] <= P * (1 - z[t])
            pr += dis[t] <= P * z[t]

    pr += pl.lpSum(sp.price_wind * (wu[t] + ws[t]) + sp.price_pv * (vu[t] + ps[t])
                   + buy_p[t] * by[t] - sp.price_sell * (ws[t] + ps[t]) for t in range(T))
    pr.solve(pl.PULP_CBC_CMD(msg=0, timeLimit=time_limit))

    r = OpResult(park=park, label="LP最优", P=P, E=E,
                 engine="MILP" if milp else "LP", status=pl.LpStatus[pr.status])
    if pr.status != pl.LpStatusOptimal:
        return r
    r.wu = np.array([wu[t].value() or 0.0 for t in range(T)])
    r.vu = np.array([vu[t].value() or 0.0 for t in range(T)])
    r.ws = np.array([ws[t].value() or 0.0 for t in range(T)])
    r.ps = np.array([ps[t].value() or 0.0 for t in range(T)])
    r.buy = np.array([by[t].value() or 0.0 for t in range(T)])
    r.ch = np.array([ch[t].value() or 0.0 for t in range(T)])
    r.dis = np.array([dis[t].value() or 0.0 for t in range(T)])
    r.soc = np.array([Ev[t].value() or 0.0 for t in range(T)])
    r.e0 = E0.value() or 0.0
    nboth = int(((r.ch > 1e-6) & (r.dis > 1e-6)).sum())
    r.status += "｜互补约束满足" if nboth == 0 else f"｜同时充放 {nboth} 时段"
    return r


def rule_strategy(L, W, V, P: float, E: float, sp: Spec, park: str = "",
                  n_days: int = 60) -> OpResult:
    """
    人工规则策略（3.3 节第 5 条）。规则：
      R1 绿电优先直供，多电源按边际成本由低到高（光伏 0.4 < 风电 0.5），弃电从贵电源起。
      R2 缺口时若 SOC 高于下限则放电：min(P, 缺口, (SOC−SOCmin)·E·ηd/Δt)。
      R3 盈余时若 SOC 低于上限则充电：min(P, 盈余, (SOCmax·E−SOC)·ηc/ηd/Δt)，优先取光伏。
      R4 其余缺口购电；盈余在允许售电时外送，否则弃电。
    为获得典型日逐日重复下的周期稳态，连续模拟 n_days 天至日初 SOC 收敛。
    """
    L = np.asarray(L, float); W = np.asarray(W, float); V = np.asarray(V, float)
    T = len(L)
    lo, hi = sp.soc_min * E, sp.soc_max * E
    soc = lo
    last = None
    wu = vu = ws = ps = by = ch = dis = np.zeros(T)
    for _ in range(max(1, n_days)):
        wu = np.zeros(T); vu = np.zeros(T); ws = np.zeros(T); ps = np.zeros(T)
        by = np.zeros(T); ch = np.zeros(T); dis = np.zeros(T)
        e = soc
        for t in range(T):
            g_w, g_v, need = W[t], V[t], L[t]
            uv = min(g_v, need); need -= uv
            uw = min(g_w, need); need -= uw
            surplus_v = g_v - uv
            surplus_w = g_w - uw
            if need > 1e-9 and e > lo + 1e-9:
                d = min(P, need, (e - lo) * sp.eta_d / sp.dt)
                dis[t] = d; need -= d; e -= d / sp.eta_d
            surplus = surplus_w + surplus_v
            if surplus > 1e-9 and e < hi - 1e-9:
                c = min(P, surplus, (hi - e) * sp.eta_c / sp.eta_d / sp.dt)
                ch[t] = c
                tv = min(surplus_v, c); surplus_v -= tv; uv += tv
                tw = min(c - tv, surplus_w); surplus_w -= tw; uw += tw
                e += sp.eta_c * c * sp.dt
            by[t] = max(need, 0.0)
            wu[t], vu[t] = uw, uv
            if sp.price_sell > 0:               # 剩余盈余外送，否则弃电
                ws[t], ps[t] = max(surplus_w, 0.0), max(surplus_v, 0.0)
        soc = e
        if last is not None and abs(e - last) < 1e-7:
            break
        last = e
    r = OpResult(park=park, label="规则策略", P=P, E=E,
                 engine="RULE", status="规则策略·周期稳态")
    r.wu, r.vu, r.ws, r.ps, r.buy, r.ch, r.dis = wu, vu, ws, ps, by, ch, dis
    r.soc = np.full(T, soc)
    r.e0 = soc
    return r


# ------------------------------------------------------------------ 进程内快速求解（HiGHS）
#  变量顺序按需剪裁，可精确减少规模：
#    不允许售电（默认）：[wu, vu, by, ch, dis, E] + E0            → 6T+1
#    允许售电          ：[wu, vu, ws, ps, by, ch, dis, E] + E0    → 8T+1
#  不允许售电时外送变量恒为 0，直接剔除可减少约 25% 变量与约束，网格搜索明显更快。
_ORDER_NOSELL = ("wu", "vu", "by", "ch", "dis", "E")
_ORDER_SELL = ("wu", "vu", "ws", "ps", "by", "ch", "dis", "E")


def fast_solve(L, W, V, P: float, E: float, sp: Spec) -> Optional[OpResult]:
    """与 solve_operation 同模型的 SciPy/HiGHS 进程内实现（不构造 PuLP 对象），用于网格搜索加速。"""
    if not HAS_SCIPY:
        return None
    L = np.asarray(L, float); W = np.asarray(W, float); V = np.asarray(V, float)
    T = len(L)
    sell = sp.price_sell > 1e-9
    order = _ORDER_SELL if sell else _ORDER_NOSELL
    idx = {k: o * T for o, k in enumerate(order)}
    n = len(order) * T + 1
    id_e0 = n - 1
    lo, hi = sp.soc_min * E, sp.soc_max * E

    c = np.zeros(n)
    c[idx["wu"]:idx["wu"] + T] = sp.price_wind
    c[idx["vu"]:idx["vu"] + T] = sp.price_pv
    c[idx["by"]:idx["by"] + T] = sp.buy_series(np.arange(T))
    if sell:
        c[idx["ws"]:idx["ws"] + T] = sp.price_wind - sp.price_sell
        c[idx["ps"]:idx["ps"] + T] = sp.price_pv - sp.price_sell

    # --- 等式约束：功率平衡 T 条 + SOC 递推 T 条 + 日循环 1 条 ---
    A_eq = np.zeros((2 * T + 1, n)); b_eq = np.zeros(2 * T + 1)
    for t in range(T):
        A_eq[t, idx["wu"] + t] = 1.0
        A_eq[t, idx["vu"] + t] = 1.0
        A_eq[t, idx["by"] + t] = 1.0
        A_eq[t, idx["dis"] + t] = 1.0
        A_eq[t, idx["ch"] + t] = -1.0
        b_eq[t] = L[t]
        r = T + t
        A_eq[r, idx["E"] + t] = 1.0
        A_eq[r, id_e0 if t == 0 else idx["E"] + t - 1] = -1.0
        A_eq[r, idx["ch"] + t] = -sp.eta_c * sp.dt
        A_eq[r, idx["dis"] + t] = sp.dt / sp.eta_d
    A_eq[2 * T, idx["E"] + T - 1] = 1.0
    A_eq[2 * T, id_e0] = -1.0

    b_wu = [(0.0, float(W[t])) for t in range(T)]
    b_vu = [(0.0, float(V[t])) for t in range(T)]
    b_by = [(0.0, None)] * T
    b_ch = [(0.0, float(P))] * T
    b_dis = [(0.0, float(P))] * T
    b_E = [(lo, hi)] * T
    b_e0 = [(lo, hi)]
    if sell:
        # 售电时需显式约束 wu+ws ≤ W、vu+ps ≤ V
        A_ub = np.zeros((2 * T, n)); b_ub = np.zeros(2 * T)
        for t in range(T):
            A_ub[t, idx["wu"] + t] = 1.0; A_ub[t, idx["ws"] + t] = 1.0; b_ub[t] = W[t]
            A_ub[T + t, idx["vu"] + t] = 1.0; A_ub[T + t, idx["ps"] + t] = 1.0; b_ub[T + t] = V[t]
        bounds = b_wu + b_vu + b_wu + b_vu + b_by + b_ch + b_dis + b_E + b_e0
    else:
        A_ub = b_ub = None                                  # 源上限已由变量上界保证
        bounds = b_wu + b_vu + b_by + b_ch + b_dis + b_E + b_e0

    res = _linprog(c, A_ub=A_ub, b_ub=b_ub, A_eq=A_eq, b_eq=b_eq, bounds=bounds, method="highs")
    if not res.success:
        return None
    x = res.x
    r = OpResult(P=P, E=E, engine="LP(HiGHS)", status="最优解")
    r.wu = x[idx["wu"]:idx["wu"] + T]
    r.vu = x[idx["vu"]:idx["vu"] + T]
    r.buy = x[idx["by"]:idx["by"] + T]
    r.ch = x[idx["ch"]:idx["ch"] + T]
    r.dis = x[idx["dis"]:idx["dis"] + T]
    r.soc = x[idx["E"]:idx["E"] + T]
    if sell:
        r.ws = x[idx["ws"]:idx["ws"] + T]
        r.ps = x[idx["ps"]:idx["ps"] + T]
    r.e0 = x[id_e0]
    return r


def op_cost(L, W, V, P: float, E: float, sp: Spec) -> Optional[float]:
    """给定 (P,E) 的日运行成本（元）。引擎优先级：HiGHS → PuLP/CBC → 规则策略。"""
    if P <= 1e-9 or E <= 1e-9:
        r = baseline(L, W, V, sp)
        return r.metrics(L, W, V, sp)["总供电成本"]
    probe_solvers()
    if SOLVER["scipy"]:
        r = fast_solve(L, W, V, P, E, sp)
        if r is not None:
            return r.metrics(L, W, V, sp)["总供电成本"]
    if SOLVER["cbc"]:
        try:
            r = _pulp_solve(L, W, V, P, E, sp)
            if r.status.startswith("Optimal"):
                return r.metrics(L, W, V, sp)["总供电成本"]
        except Exception as e:
            SOLVER["cbc"] = False
            SOLVER["cbc_msg"] = f"{type(e).__name__}: {str(e)[:160]}"
    # 最后手段：纯 Python 规则策略（"单一电价 + 不允许售电"下即精确最优，其余场景为近似）
    SOLVER["rule_fallback"] = True
    r = rule_strategy(L, W, V, P, E, sp)
    return r.metrics(L, W, V, sp)["总供电成本"]


def optimize_capacity(L, W, V, sp: Spec, park: str = "", P_list=None, E_list=None,
                      n_refine: int = 5, scale_cost: float = 1.0,
                      scale_price_buy: float = 1.0, eta: Optional[float] = None,
                      life_y: Optional[int] = None,
                      log: Optional[Callable[[str], None]] = None,
                      progress: Optional[Callable[[int, int, str], None]] = None,
                      verify: bool = True) -> Dict[str, Any]:
    """
    在 (P,E) 网格上最小化年化总成本（= 年折旧投资 + 年运行成本），再做迭代细化。

    verify=True 时对最优解用 PuLP 复算做交叉验证（单次容量优化建议开启）；
    敏感性分析、联合运营等批量调用应传 verify=False，否则每个配置都会多启动一次
    CBC 子进程，在无 SciPy 的机器上是主要耗时来源。
    长任务可通过 request_cancel() 中止，内部按网格点检查。
    """
    sp = sp.clone(cost_p=sp.cost_p * scale_cost, cost_e=sp.cost_e * scale_cost,
                  price_buy=sp.price_buy * scale_price_buy)
    if eta is not None:
        sp.eta_c = sp.eta_d = eta
    if life_y is not None:
        sp.life_y = life_y

    P_list = list(P_list) if P_list is not None else [0.0] + list(np.arange(10, 260.1, 10))
    E_list = list(E_list) if E_list is not None else [0.0] + list(np.arange(25, 550.1, 25))
    days = sp.days

    n_base = sum(1 for P in P_list for E in E_list if not ((P <= 0) != (E <= 0)))
    total_est = n_base + max(0, n_refine) * 121 + 2
    done = {"n": 0}

    base_cost = op_cost(L, W, V, 0.0, 0.0, sp)

    def scan(Ps, Es, best):
        surf = np.full((len(Ps), len(Es)), np.nan)
        for i, P in enumerate(Ps):
            for j, E in enumerate(Es):
                if (P <= 0) != (E <= 0):
                    continue
                check_cancel()                     # 每个网格点检查一次，可及时响应中止
                c = op_cost(L, W, V, P, E, sp)
                done["n"] += 1
                if progress is not None and done["n"] % 25 == 0:
                    progress(done["n"], total_est, park or "")
                if c is None:
                    continue
                inv = sp.cost_p * P + sp.cost_e * E
                tot = inv / sp.life_y + c * days
                surf[i, j] = tot
                if best is None or tot < best["年化总成本"] - 1e-9:
                    best = {"P": P, "E": E, "投资总额": inv, "年折旧投资": inv / sp.life_y,
                            "年运行成本": c * days, "日运行成本": c, "年化总成本": tot,
                            "全寿命总成本": inv + c * days * sp.life_y}
        return best, surf

    if log:
        log(f"容量优化：粗网格 {len(P_list)}×{len(E_list)}（有效点 {n_base}），"
            f"细化 {n_refine} 轮，预计求解约 {total_est} 次 LP…")
    best, surf = scan(P_list, E_list, None)
    if best is None:
        raise RuntimeError("容量优化未得到可行解（请检查参数或求解器）")

    refined = False
    if best["P"] > 0:
        dp = (P_list[1] - P_list[0]) if len(P_list) > 1 else 10.0
        de = (E_list[1] - E_list[0]) if len(E_list) > 1 else 25.0
        pc, ec, sp_, ep_ = best["P"], best["E"], dp / 5.0, de / 5.0
        for _ in range(max(0, n_refine)):
            Ps = sorted({round(max(1.0, pc + k * sp_), 4) for k in range(-5, 6) if pc + k * sp_ > 0})
            Es = sorted({round(max(1.0, ec + k * ep_), 4) for k in range(-5, 6) if ec + k * ep_ > 0})
            best2, _ = scan(Ps, Es, dict(best))
            if best2["年化总成本"] < best["年化总成本"] - 1e-7:
                best, pc, ec = best2, best2["P"], best2["E"]
                refined = True
            sp_, ep_ = sp_ / 2.0, ep_ / 2.0
        if base_cost * days < best["年化总成本"] - 1e-9:
            best = {"P": 0.0, "E": 0.0, "投资总额": 0.0, "年折旧投资": 0.0,
                    "年运行成本": base_cost * days, "日运行成本": base_cost,
                    "年化总成本": base_cost * days,
                    "全寿命总成本": base_cost * days * sp.life_y}
            refined = True
    best["细化搜索"] = refined

    # 交叉验证（批量调用时关闭，避免反复启动 CBC 子进程）；无 CBC 时改用双引擎互校
    probe_solvers()
    if verify and best["P"] > 0:
        if SOLVER["cbc"]:
            try:
                rp = _pulp_solve(L, W, V, best["P"], best["E"], sp, park=park)
                if rp.status.startswith("Optimal"):
                    cv = rp.metrics(L, W, V, sp)["总供电成本"]
                    best["复算日成本"] = cv
                    best["求解器偏差"] = abs(cv - best["日运行成本"])
                    best["交叉验证"] = "HiGHS ↔ PuLP/CBC"
                    best["result"] = rp
            except Exception as e:
                SOLVER["cbc"] = False
                SOLVER["cbc_msg"] = f"{type(e).__name__}: {str(e)[:160]}"
                best["交叉验证"] = "跳过（CBC 不可用）"
        elif SOLVER["scipy"]:
            r2 = fast_solve(L, W, V, best["P"], best["E"], sp)
            r3 = rule_strategy(L, W, V, best["P"], best["E"], sp)
            if r2 is not None:
                c2 = r2.metrics(L, W, V, sp)["总供电成本"]
                c3 = r3.metrics(L, W, V, sp)["总供电成本"]
                best["复算日成本"] = c2
                best["求解器偏差"] = abs(c3 - c2)
                best["交叉验证"] = "HiGHS ↔ 规则策略"
                best["result"] = r2
        else:
            best["交叉验证"] = "跳过（无可用 LP 求解器）"
    return {"best": best, "surf": surf, "P_list": P_list, "E_list": E_list,
            "spec": sp, "base_cost": base_cost}


# ------------------------------------------------------------------ 网格生成与工作量预估
#  在"出结果"与"等多久"之间做取舍：界面按精度档位生成网格，并预估总耗时后再开跑。
def build_grid(step_p: float, step_e: float, pmax: float,
               emax: float) -> Tuple[List[float], List[float]]:
    """按步长与上限生成含 0 点的 (P_list, E_list)。"""
    Pl = [0.0] + [round(float(x), 2) for x in np.arange(step_p, pmax + 1e-9, step_p)]
    El = [0.0] + [round(float(x), 2) for x in np.arange(step_e, emax + 1e-9, step_e)]
    return Pl, El


def count_grid_points(Pl: Sequence[float], El: Sequence[float]) -> int:
    """有效网格点数（排除 (P=0,E>0) 与 (P>0,E=0) 两种非法组合）。"""
    return sum(1 for P in Pl for E in El if not ((P <= 0) != (E <= 0)))


def estimate_solves(Pl: Sequence[float], El: Sequence[float], n_refine: int) -> int:
    return count_grid_points(Pl, El) + max(0, n_refine) * 121


def per_solve_seconds() -> float:
    """单次 LP 的粗略耗时（秒）：SciPy/HiGHS 进程内约 3 ms；PuLP 需启动 CBC 子进程，慢约两个数量级。"""
    return 0.003 if HAS_SCIPY else 0.35


def solve_budget_note(n_solves: int, n_cfg: int = 1) -> str:
    sec = n_solves * n_cfg * per_solve_seconds()
    return (f"约 {n_solves} 次 LP/配置 × {n_cfg} 配置 = {n_solves * n_cfg} 次，"
            f"预计 {sec:.0f} 秒" + ("" if HAS_SCIPY else "（未检测到 SciPy，按 PuLP 慢速估算）"))


def solve_joint(data: Data, P: float, E: float, sp: Spec, month: int = 0,
                milp: bool = False) -> OpResult:
    L, W, V = data.joint_arrays(month)
    return solve_operation(L, W, V, P, E, sp, park="JOINT", milp=milp)


# ==============================================================================
#  §5  结果自检（物理可行性 + 双求解器交叉验证）
# ==============================================================================
def validate_dispatch(r: OpResult, L, W, V, sp: Spec) -> List[Dict[str, Any]]:
    L = np.asarray(L, float); W = np.asarray(W, float); V = np.asarray(V, float)
    out: List[Dict[str, Any]] = []

    def add(name, val, tol, ok):
        out.append({"检查项": name, "实测": val, "容差": tol, "结论": "通过" if ok else "未通过"})

    bal = np.max(np.abs(r.wu + r.vu + r.dis + r.buy - L - r.ch)) if len(L) else 0.0
    add("功率平衡残差 (kW)", fnum(bal, 6), "<1e-4", bal < 1e-4)
    ov = max(float(np.max(r.wu + r.ws - W)) if len(W) else 0.0,
             float(np.max(r.vu + r.ps - V)) if len(V) else 0.0)
    add("绿电不超装机 (kW)", fnum(ov, 6), "≤1e-4", ov <= 1e-4)
    neg = min(float(np.min(r.wu)), float(np.min(r.vu)), float(np.min(r.buy)),
              float(np.min(r.ch)), float(np.min(r.dis)))
    add("变量非负性", fnum(neg, 6), "≥-1e-6", neg > -1e-6)
    if r.P > 0 and r.E > 0 and len(r.soc) and float(np.max(r.soc)) > 0:
        smin, smax = sp.soc_min * r.E, sp.soc_max * r.E
        viol = max(max(0.0, smin - float(np.min(r.soc))), max(0.0, float(np.max(r.soc)) - smax))
        add("SOC 越限 (kWh)", fnum(viol, 6), "<1e-4", viol < 1e-4)
        pmax = max(float(np.max(r.ch)), float(np.max(r.dis)))
        add("充放功率越限 (kW)", fnum(max(0.0, pmax - r.P), 6), "<1e-4", pmax <= r.P + 1e-4)
        nboth = int(((r.ch > 1e-6) & (r.dis > 1e-6)).sum())
        add("充放电互斥 (时段数)", str(nboth), "=0", nboth == 0)
        if r.engine.startswith("LP") or r.engine == "MILP":
            cyc = abs(float(r.soc[-1]) - r.e0)
            add("日循环稳态 |E未−E初| (kWh)", fnum(cyc, 6), "<1e-4", cyc < 1e-4)
    else:
        add("储能约束", "无储能", "—", True)
    return out


# ==============================================================================
#  §6  报表导出
# ==============================================================================
class ReportWriter:
    """把当前结果写为 CSV / PNG / TXT。"""

    def __init__(self, outdir: str, stamp: str = ""):
        self.outdir = outdir
        self.stamp = stamp or _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        os.makedirs(outdir, exist_ok=True)

    def path(self, stem: str, ext: str) -> str:
        return _osp.join(self.outdir, f"{stem}_{self.stamp}.{ext}")

    def write_csv(self, stem: str, rows: Sequence[Dict[str, Any]]) -> Optional[str]:
        if not rows:
            return None
        p = self.path(stem, "csv")
        keys: List[str] = []
        for r in rows:
            for k in r:
                if k not in keys:
                    keys.append(k)
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=keys)
            w.writeheader()
            for r in rows:
                w.writerow({k: r.get(k, "") for k in keys})
        return p

    def write_png(self, fig) -> Optional[str]:
        if not HAS_MPL or fig is None:
            return None
        p = self.path("figure", "png")
        fig.savefig(p, dpi=160, bbox_inches="tight")
        return p

    def write_text(self, stem: str, lines: Sequence[str]) -> str:
        p = self.path(stem, "txt")
        with open(p, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        return p


# ==============================================================================
#  §7  GUI
# ==============================================================================
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QIcon, QTextCursor
from PyQt5.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDoubleSpinBox, QFileDialog,
    QGridLayout, QGroupBox, QHBoxLayout, QHeaderView, QLabel, QMainWindow, QMessageBox,
    QPlainTextEdit, QProgressBar, QPushButton, QScrollArea, QSizePolicy, QSpinBox, QSplitter,
    QStatusBar,
    QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget,
)

X24 = np.arange(24)


class Worker(QThread):
    """后台执行耗时任务（容量优化 / 联合运营 / 敏感性分析），支持进度回报与中止。"""
    done = pyqtSignal(dict)
    fail = pyqtSignal(str)
    stopped = pyqtSignal()
    log = pyqtSignal(str)
    prog = pyqtSignal(int, int, str)

    def __init__(self, fn: Callable[..., Dict[str, Any]], parent=None):
        super().__init__(parent)
        self.fn = fn

    def run(self) -> None:  # noqa: D401
        try:
            self.done.emit(self.fn(self.log.emit, self.prog.emit))
        except Cancelled:
            self.stopped.emit()
        except Exception as e:  # pragma: no cover
            self.fail.emit(f"{type(e).__name__}: {e}\n{traceback.format_exc()}")


class SelfTestRunner:
    """内置回归自检：使用题目默认参数，独立于界面上的当前参数，逐项验证模型与论文一致性。"""

    def __init__(self, data: Data):
        self.data = data
        self.rows: List[Dict[str, Any]] = []

    def _rec(self, name: str, ok: bool, detail: str = "") -> bool:
        self.rows.append({"检查项": name, "结论": "通过" if ok else "未通过", "说明": detail})
        return ok

    def run(self, log: Callable[[str], None]) -> List[Dict[str, Any]]:
        d, sp = self.data, Spec()          # 固定题目默认参数

        ok = len(d.hours) == 24 and all(len(d.load[p]) == 24 for p in PARKS)
        self._rec("数据完整性（24 时段、无缺失）", ok, f"{len(d.hours)} 时段，三园区负荷轴一致")

        exp = {"A": 0.7701, "B": 0.6577, "C": 0.6383}
        got = {}
        for p in PARKS:
            L, W, V = d.arrays(p, 0)
            got[p] = baseline(L, W, V, sp, p).metrics(L, W, V, sp)["单位电量平均供电成本"]
        self._rec("无储能基线复现论文（0.7701/0.6577/0.6383）",
                  all(abs(got[k] - exp[k]) < 1e-3 for k in exp),
                  " / ".join(f"{k}={got[k]:.4f}" for k in exp))

        L, W, V = d.arrays("A", 0)
        b1 = baseline(L, W, V, sp, "A").metrics(L, W, V, sp)["总供电成本"]
        b2 = baseline(L, W, V, sp, "A").metrics(L, W, V, sp)["总供电成本"]
        self._rec("基线确定性（重复计算一致，可复核）", abs(b1 - b2) < 1e-12, f"{b1:.6f} 元/日")

        L, W, V = d.arrays("B", 0)
        r0 = baseline(L, W, V, sp, "B")
        r1 = rule_strategy(L, W, V, 50, 100, sp, "B")
        r2 = solve_operation(L, W, V, 50, 100, sp, "B")
        c0 = r0.metrics(L, W, V, sp)["总供电成本"]
        c1 = r1.metrics(L, W, V, sp)["总供电成本"]
        c2 = r2.metrics(L, W, V, sp)["总供电成本"]
        self._rec("LP 最优不劣于人工规则策略", c2 <= c1 + 1e-6,
                  f"规则 {c1:.3f} vs LP {c2:.3f} 元/日（差 {c1 - c2:.4f}）")
        self._rec("LP 最优不劣于无储能基线", c2 <= c0 + 1e-6, f"基线 {c0:.3f} 元/日")
        self._rec("LP 充放电互斥（互补）满足", "互补约束满足" in r2.status, r2.status)

        if HAS_SCIPY:
            rs = fast_solve(L, W, V, 50, 100, sp)
            dv = abs(rs.metrics(L, W, V, sp)["总供电成本"] - c2) if rs else 9e9
            self._rec("PuLP 与 SciPy/HiGHS 双求解器交叉验证", dv < 1e-4, f"偏差 {dv:.2e} 元")

        L, W, V = d.arrays("C", 0)
        rl = solve_operation(L, W, V, 50, 100, sp, "C")
        rm = solve_operation(L, W, V, 50, 100, sp, "C", milp=True)
        dv = abs(rl.metrics(L, W, V, sp)["总供电成本"] - rm.metrics(L, W, V, sp)["总供电成本"])
        self._rec("MILP（硬互斥）与 LP 目标一致", dv < 1e-3, f"偏差 {dv:.2e} 元")

        log("回归自检：容量优化复现中（A/B/C）…")
        Pl = [0.0] + [round(x, 2) for x in np.arange(10, 150.1, 10)]
        El = [0.0] + [round(x, 2) for x in np.arange(25, 400.1, 25)]
        best = {}
        for p in PARKS:
            L, W, V = d.arrays(p, 0)
            best[p] = optimize_capacity(L, W, V, sp, park=p, P_list=Pl, E_list=El, n_refine=3)["best"]
        self._rec("园区A 最优 = 不配置储能", best["A"]["P"] == 0,
                  f"P={best['A']['P']:.0f} kW（50/100 年亏 6554.7 元）")
        self._rec("园区B 最优 ≈ 125 kW / 353.75 kWh",
                  abs(best["B"]["P"] - 125) < 5 and abs(best["B"]["E"] - 353.75) < 10,
                  f"{best['B']['P']:.1f}/{best['B']['E']:.1f}")
        self._rec("园区C 最优 ≈ 58.75 kW / 70 kWh",
                  abs(best["C"]["P"] - 58.75) < 5 and abs(best["C"]["E"] - 70) < 10,
                  f"{best['C']['P']:.1f}/{best['C']['E']:.1f}")

        ind = sum(op_cost(*d.arrays(p, 0), 0, 0, sp) for p in PARKS)
        jn = op_cost(*d.joint_arrays(0), 0, 0, sp)
        self._rec("联合运营优于独立运营（无储能）", jn < ind,
                  f"独立 {ind:.1f} → 联合 {jn:.1f} 元/日（省 {ind - jn:.1f}）")

        L, W, V = d.arrays("C", 0)
        vr = validate_dispatch(solve_operation(L, W, V, 50, 100, sp, "C"), L, W, V, sp)
        np_ = sum(1 for x in vr if x["结论"] == "通过")
        self._rec("调度物理可行性（平衡/SOC/功率/非负）", np_ == len(vr), f"{np_}/{len(vr)} 项")

        spS = sp.clone(price_sell=0.45)
        ms = solve_operation(L, W, V, 50, 100, spS, "C").metrics(L, W, V, spS)
        self._rec("售电价高于绿电成本 → 售电生效", ms["售电量"] > 1e-6,
                  f"售电 {ms['售电量']:.1f} kWh，收益 {ms['售电收益']:.1f} 元")
        spL = sp.clone(price_sell=0.30)
        ml = solve_operation(L, W, V, 50, 100, spL, "C").metrics(L, W, V, spL)
        self._rec("售电价低于绿电成本 → 不售电（经济正确）", abs(ml["售电量"]) < 1e-6,
                  f"售电 {ml['售电量']:.6f} kWh")

        spT = sp.clone(tou_enable=True)
        L, W, V = d.arrays("B", 0)
        ct2 = solve_operation(L, W, V, 50, 100, spT, "B").metrics(L, W, V, spT)["总供电成本"]
        ct1 = rule_strategy(L, W, V, 50, 100, spT, "B").metrics(L, W, V, spT)["总供电成本"]
        self._rec("分时电价下 LP 优于人工规则", ct2 <= ct1 + 1e-6,
                  f"规则 {ct1:.2f} vs LP {ct2:.2f} 元/日")
        return self.rows


class MainWindow(QMainWindow):
    """主界面：顶部操作按钮、左侧参数面板、右侧图形与数据表页签。"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION}")
        self.resize(1580, 960)
        self.data: Optional[Data] = None
        self.view = "welcome"
        self.payload: Dict[str, Any] = {}
        self.worker: Optional[QThread] = None
        self._build_ui()
        self._log(f"{APP_NAME} {APP_VERSION} 已就绪")
        self._log(f"Python {sys.version.split()[0]}｜"
                  f"绘图后端：{'Matplotlib' if HAS_MPL else '未安装 matplotlib（图形区不可用）'}")
        for ln in engine_report():            # 求解器自检：实测各引擎是否真的可用
            self._log(ln)
        self._apply_solver_capability()
        self._log("操作提示：先点「① 数据检查」，再按 ②→⑦ 逐步运行；所有参数改动即时生效。")
        self._auto_load()

    # ------------------------------------------------------------------ 界面
    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(8, 6, 8, 6)
        root.setSpacing(6)
        root.addWidget(self._build_toolbar())
        split = QSplitter(Qt.Horizontal)
        split.addWidget(self._build_param_panel())
        split.addWidget(self._build_right_panel())
        split.setStretchFactor(0, 0)
        split.setStretchFactor(1, 1)
        split.setSizes([400, 1180])
        root.addWidget(split, 1)
        self.status = QStatusBar()
        self.progress = QProgressBar()
        self.progress.setMaximumWidth(240)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setVisible(False)
        self.status.addPermanentWidget(self.progress)
        self.setStatusBar(self.status)
        self.status.showMessage("就绪")
        self._last_pct = -1

    def _build_toolbar(self) -> QWidget:
        box = QGroupBox("操作")
        box.setStyleSheet("QGroupBox { font-weight: bold; }")
        lay = QHBoxLayout(box)
        lay.setContentsMargins(10, 6, 10, 6)
        lay.setSpacing(6)

        def mk(text: str, color: str, slot) -> QPushButton:
            b = QPushButton(text)
            b.setMinimumHeight(32)
            b.setStyleSheet(
                f"QPushButton {{ background-color: {color}; color: white; font-weight: bold;"
                f" border-radius: 5px; padding: 4px 10px; }}"
                f" QPushButton:disabled {{ background-color: #9e9e9e; }}")
            b.clicked.connect(slot)
            return b

        self.btn_data = mk("① 数据检查", "#2E86AB", self.on_data)
        self.btn_base = mk("② 无储能基线", "#1F6FB4", self.on_baseline)
        self.btn_sim = mk("③ 储能运行模拟", "#1a7f37", self.on_simulate)
        self.btn_opt = mk("④ 最优运行策略", "#0f7b6c", self.on_optimal)
        self.btn_cap = mk("⑤ 容量优化", "#8a6d00", self.on_capacity)
        self.btn_joint = mk("⑥ 联合运营", "#6a3fa0", self.on_joint)
        self.btn_sens = mk("⑦ 敏感性分析", "#b35c00", self.on_sensitivity)
        self.btn_val = mk("⑧ 结果自检", "#0b6e99", self.on_validate)
        self.btn_self = mk("⑨ 一键回归自检", "#6a3fa0", self.on_selftest)
        self.btn_exp = mk("导出报表", "#5d4037", self.on_export)
        self.btn_stop = mk("■ 停止计算", "#b00020", self.on_stop)
        self.btn_stop.setEnabled(False)
        self.btn_clr = mk("清空", "#616161", self.on_clear)
        self.buttons = [self.btn_data, self.btn_base, self.btn_sim, self.btn_opt,
                        self.btn_cap, self.btn_joint, self.btn_sens, self.btn_val,
                        self.btn_self, self.btn_exp]
        for b in self.buttons:
            lay.addWidget(b)
        lay.addWidget(self.btn_stop)          # 停止键不随 _busy 一起禁用，运行中始终可点
        lay.addWidget(self.btn_clr)
        lay.addStretch()
        self.lbl_engine = QLabel("")
        self.lbl_engine.setStyleSheet("color:#1a7f37; font-weight:bold;")
        lay.addWidget(self.lbl_engine)
        return box

    def _update_engine_label(self) -> None:
        self.lbl_engine.setText("主求解引擎：" + primary_engine())

    def _apply_solver_capability(self) -> None:
        """按实测的求解器能力调整界面（MILP 硬约束依赖 CBC，不可用时自动禁用）。"""
        probe_solvers()
        self._update_engine_label()
        if not SOLVER["cbc"]:
            self.chk_milp.blockSignals(True)
            self.chk_milp.setChecked(False)
            self.chk_milp.blockSignals(False)
            self.chk_milp.setEnabled(False)
            self.chk_milp.setToolTip(
                "本机 PuLP/CBC 求解器无法启动，已自动禁用 MILP 硬约束。\n"
                "LP（HiGHS）解本身已满足充放电互斥，结果不受影响。\n"
                "如需启用：安装 Microsoft Visual C++ 2015-2022 (x64) 运行库，"
                "并把 %TEMP% 加入杀毒软件白名单。")
        else:
            self.chk_milp.setEnabled(True)

    # ------------------------------------------------------------ 参数面板
    def _spin(self, lo, hi, val, dec=4, step=1.0) -> QDoubleSpinBox:
        s = QDoubleSpinBox()
        s.setRange(lo, hi)
        s.setDecimals(dec)
        s.setSingleStep(step)
        s.setValue(val)
        s.setMinimumWidth(96)
        s.setMaximumWidth(140)
        return s

    def _ispin(self, lo, hi, val) -> QSpinBox:
        s = QSpinBox()
        s.setRange(lo, hi)
        s.setValue(val)
        s.setMinimumWidth(96)
        s.setMaximumWidth(140)
        return s

    def _row(self, grid: QGridLayout, r: int, label: str, widget: QWidget, hint: str = "") -> int:
        lab = QLabel(label + ("  ⓘ" if hint else ""))
        if hint:
            lab.setToolTip(hint)
        grid.addWidget(lab, r, 0)
        grid.addWidget(widget, r, 1)
        return r + 1

    def _build_param_panel(self) -> QWidget:
        holder = QWidget()
        outer = QVBoxLayout(holder)
        outer.setContentsMargins(0, 0, 4, 0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setMinimumWidth(400)
        scroll.setMaximumWidth(480)
        inner = QWidget()
        v = QVBoxLayout(inner)
        v.setContentsMargins(2, 2, 2, 2)
        v.setSpacing(8)

        # ---- ① 数据与园区 ----
        g = QGroupBox("① 数据与园区")
        grid = QGridLayout(g)
        r = 0
        self.cb_park = QComboBox()
        for p in PARKS:
            self.cb_park.addItem(f"园区 {p}", p)
        self.cb_park.addItem("三园区联合", "JOINT")
        r = self._row(grid, r, "分析对象", self.cb_park, "选择单个园区或三园区联合运营")
        self.cb_src = QComboBox()
        self.cb_src.addItem("典型日（附件2）", 0)
        for m in range(1, 13):
            self.cb_src.addItem(f"{m} 月（附件3）", m)
        r = self._row(grid, r, "数据时段", self.cb_src,
                      "典型日=附件2；1~12 月=附件3，可逐月验证设想")
        self.sp_loadscale = self._spin(10, 300, 100.0, 1, 5.0)
        r = self._row(grid, r, "负荷缩放 (%)", self.sp_loadscale,
                      "按比例缩放三园区负荷，用于验证负荷增长/下降情景")
        self.sp_windcap = self._spin(0, 100000, 1000.0, 1, 50.0)
        r = self._row(grid, r, "风电装机 (kW)", self.sp_windcap,
                      "所选园区的风电装机；该园区无风电时不可编辑（等价于调整风电出力）")
        self.sp_pvcap = self._spin(0, 100000, 750.0, 1, 50.0)
        r = self._row(grid, r, "光伏装机 (kW)", self.sp_pvcap,
                      "所选园区的光伏装机；该园区无光伏时不可编辑（等价于调整光伏出力）")
        self.btn_xl = QPushButton("载入 Excel…")
        self.btn_xl.setMinimumHeight(28)
        self.btn_xl.setToolTip("选择自备的附件 Excel（可多选，按文件名自动识别附件1/2/3）")
        self.btn_xl.clicked.connect(self.on_load_excel)
        self.btn_emb = QPushButton("恢复内置")
        self.btn_emb.setMinimumHeight(28)
        self.btn_emb.setToolTip("切回代码内置的数据模板")
        self.btn_emb.clicked.connect(self.on_use_embedded)
        grid.addWidget(self.btn_xl, r, 0)
        grid.addWidget(self.btn_emb, r, 1)
        r += 1
        self.lbl_src = QLabel("当前：—")
        self.lbl_src.setWordWrap(True)
        self.lbl_src.setStyleSheet("color:#555;")
        grid.addWidget(self.lbl_src, r, 0, 1, 2)
        r += 1
        v.addWidget(g)

        # ---- ② 价格与成本 ----
        g = QGroupBox("② 价格与成本（元/kWh）")
        grid = QGridLayout(g)
        r = 0
        self.sp_buy = self._spin(0, 10, 1.0, 4, 0.1)
        r = self._row(grid, r, "购电电价", self.sp_buy, "题目给定 1.0")
        self.sp_wind = self._spin(0, 10, 0.5, 4, 0.05)
        r = self._row(grid, r, "风电使用成本", self.sp_wind)
        self.sp_pv = self._spin(0, 10, 0.4, 4, 0.05)
        r = self._row(grid, r, "光伏使用成本", self.sp_pv)
        self.chk_sell = QCheckBox("允许富余绿电上网售电")
        self.chk_sell.setToolTip("题目设定为不允许上网；勾选后可验证'放开售电'设想")
        self.chk_sell.stateChanged.connect(self._sync_sell)
        grid.addWidget(self.chk_sell, r, 0, 1, 2); r += 1
        self.sp_sell = self._spin(0, 10, 0.3, 4, 0.05)
        self.sp_sell.setEnabled(False)
        r = self._row(grid, r, "售电电价", self.sp_sell, "仅勾选'允许售电'时生效")
        self.chk_tou = QCheckBox("启用分时电价")
        self.chk_tou.setToolTip("题目为单一电价；启用后可验证 LP 相对规则策略的优势")
        self.chk_tou.stateChanged.connect(self._sync_tou)
        grid.addWidget(self.chk_tou, r, 0, 1, 2); r += 1
        self.sp_peak = self._spin(0, 10, 1.2, 4, 0.1)
        r = self._row(grid, r, "峰价 (8-11,18-21 时)", self.sp_peak)
        self.sp_flat = self._spin(0, 10, 1.0, 4, 0.1)
        r = self._row(grid, r, "平价 (其余时段)", self.sp_flat)
        self.sp_valley = self._spin(0, 10, 0.6, 4, 0.1)
        r = self._row(grid, r, "谷价 (23-6 时)", self.sp_valley)
        for w in (self.sp_peak, self.sp_flat, self.sp_valley):
            w.setEnabled(False)
        v.addWidget(g)

        # ---- ③ 储能参数 ----
        g = QGroupBox("③ 储能参数")
        grid = QGridLayout(g)
        r = 0
        self.sp_P = self._spin(0, 5000, 50.0, 2, 10.0)
        r = self._row(grid, r, "储能功率 P (kW)", self.sp_P, "题目给定方案 50 kW")
        self.sp_E = self._spin(0, 20000, 100.0, 2, 10.0)
        r = self._row(grid, r, "储能容量 E (kWh)", self.sp_E, "题目给定方案 100 kWh")
        self.sp_cp = self._spin(0, 100000, 800.0, 2, 50.0)
        r = self._row(grid, r, "功率单价 (元/kW)", self.sp_cp)
        self.sp_ce = self._spin(0, 100000, 1800.0, 2, 100.0)
        r = self._row(grid, r, "容量单价 (元/kWh)", self.sp_ce)
        self.sp_etac = self._spin(0.05, 1.0, 0.95, 4, 0.01)
        r = self._row(grid, r, "充电效率 ηc", self.sp_etac, "题目给定 0.95")
        self.sp_etad = self._spin(0.05, 1.0, 0.95, 4, 0.01)
        r = self._row(grid, r, "放电效率 ηd", self.sp_etad)
        self.sp_smin = self._spin(0, 0.9, 0.10, 4, 0.05)
        r = self._row(grid, r, "SOC 下限", self.sp_smin)
        self.sp_smax = self._spin(0.1, 1.0, 0.90, 4, 0.05)
        r = self._row(grid, r, "SOC 上限", self.sp_smax)
        self.sp_life = self._ispin(1, 40, 10)
        r = self._row(grid, r, "运行寿命 (年)", self.sp_life)
        self.sp_days = self._ispin(1, 366, 365)
        r = self._row(grid, r, "年运行天数", self.sp_days)
        v.addWidget(g)

        # ---- ④ 优化设置 ----
        g = QGroupBox("④ 优化设置")
        grid = QGridLayout(g)
        r = 0
        self.chk_milp = QCheckBox("MILP 硬约束（充放互斥）")
        self.chk_milp.setToolTip("LP 松弛通常已满足互斥；勾选后引入 0-1 变量强制互斥，求解更慢")
        grid.addWidget(self.chk_milp, r, 0, 1, 2); r += 1
        self.sp_tl = self._ispin(1, 600, 30)
        r = self._row(grid, r, "求解时限 (s)", self.sp_tl)
        self.sp_pmax = self._spin(0, 2000, 260.0, 1, 10.0)
        r = self._row(grid, r, "优化 P 上限 (kW)", self.sp_pmax)
        self.sp_pstep = self._spin(1, 200, 10.0, 2, 1.0)
        r = self._row(grid, r, "P 步长 (kW)", self.sp_pstep)
        self.sp_emax = self._spin(0, 5000, 550.0, 1, 25.0)
        r = self._row(grid, r, "优化 E 上限 (kWh)", self.sp_emax)
        self.sp_estep = self._spin(1, 500, 25.0, 2, 1.0)
        r = self._row(grid, r, "E 步长 (kWh)", self.sp_estep)
        self.sp_refine = self._ispin(0, 8, 4)
        r = self._row(grid, r, "细化轮数", self.sp_refine, "在最优解附近逐轮减半步长再搜索")
        self.cb_sens_prec = QComboBox()
        self.cb_sens_prec.addItem("粗（最快，约 5~10 秒）", "coarse")
        self.cb_sens_prec.addItem("中（推荐，约 30 秒）", "mid")
        self.cb_sens_prec.addItem("细（最准，约 2~3 分钟）", "fine")
        self.cb_sens_prec.setCurrentIndex(1)
        r = self._row(grid, r, "敏感性精度", self.cb_sens_prec,
                      "敏感性分析对每个参数档都要重做一次容量优化；精度越高耗时越长。"
                      "未安装 SciPy 时会自动放粗网格并给出提示。")
        v.addWidget(g)

        note = QLabel("提示：默认参数严格复现论文结果；改动任一参数后重跑对应步骤即可对比。"
                      "数据默认使用代码内置模板，可用「载入 Excel 数据」替换。")
        note.setWordWrap(True)
        note.setStyleSheet("color:#555; padding:4px;")
        v.addWidget(note)
        v.addStretch()

        # 联动在全部控件创建完成后连接，避免构建期触发回调
        self.cb_park.currentIndexChanged.connect(self._sync_park)
        self.sp_windcap.valueChanged.connect(self._on_cap_changed)
        self.sp_pvcap.valueChanged.connect(self._on_cap_changed)

        scroll.setWidget(inner)
        outer.addWidget(scroll)
        return holder

    def _sync_sell(self) -> None:
        self.sp_sell.setEnabled(self.chk_sell.isChecked())

    def _sync_tou(self) -> None:
        on = self.chk_tou.isChecked()
        for w in (self.sp_peak, self.sp_flat, self.sp_valley):
            w.setEnabled(on)

    # ------------------------------------------------------------ 右侧面板
    def _build_right_panel(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)

        if HAS_MPL:
            self.fig = Figure(figsize=(10, 6.4), dpi=100)
            self.canvas = FigureCanvas(self.fig)
            self.canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self.tb = NavToolbar(self.canvas, self)
            lay.addWidget(self.tb)
            lay.addWidget(self.canvas, 3)
        else:
            self.fig = None
            self.canvas = None
            lay.addWidget(QLabel("未安装 matplotlib，图形区不可用（仅表格可用）"), 3)

        self.tabs = QTabWidget()
        self.tab_log = QPlainTextEdit()
        self.tab_log.setReadOnly(True)
        self.tabs.addTab(self.tab_log, "运行日志")
        self.tables: Dict[str, QTableWidget] = {}
        for name in ("数据检查", "无储能基线", "方案对比", "容量优化", "联合运营", "敏感性", "结果自检", "调度明细"):
            t = self._mk_table()
            self.tables[name] = t
            self.tabs.addTab(t, name)
        lay.addWidget(self.tabs, 4)
        return w

    def _mk_table(self) -> QTableWidget:
        t = QTableWidget()
        t.setEditTriggers(QAbstractItemView.NoEditTriggers)
        t.setAlternatingRowColors(True)
        t.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        t.verticalHeader().setDefaultSectionSize(22)
        return t

    # ------------------------------------------------------------------ 参数收集
    def _spec(self) -> Spec:
        return Spec(
            price_buy=self.sp_buy.value(), price_wind=self.sp_wind.value(),
            price_pv=self.sp_pv.value(),
            price_sell=(self.sp_sell.value() if self.chk_sell.isChecked() else 0.0),
            tou_enable=self.chk_tou.isChecked(), tou_peak=self.sp_peak.value(),
            tou_flat=self.sp_flat.value(), tou_valley=self.sp_valley.value(),
            cost_p=self.sp_cp.value(), cost_e=self.sp_ce.value(),
            eta_c=self.sp_etac.value(), eta_d=self.sp_etad.value(),
            soc_min=self.sp_smin.value(), soc_max=self.sp_smax.value(),
            life_y=self.sp_life.value(), days=self.sp_days.value(),
        )

    def _cur_month(self) -> int:
        return int(self.cb_src.currentData())

    def _cur_park(self) -> str:
        return str(self.cb_park.currentData())

    def _arrays(self, park: str, month: int):
        """取得 (负荷, 风电, 光伏) kW，并应用界面上的负荷缩放系数。"""
        if park == "JOINT":
            L, W, V = self.data.joint_arrays(month)
        else:
            L, W, V = self.data.arrays(park, month)
        return L * (self.sp_loadscale.value() / 100.0), W, V

    # ------------------------------------------------------------ 数据来源联动
    def _sync_park(self) -> None:
        """切换分析对象时，把所选园区的装机填入输入框；无该电源则禁用。"""
        park = self._cur_park()
        for w, kind in ((self.sp_windcap, "wind"), (self.sp_pvcap, "pv")):
            has = (park != "JOINT") and (PU_SRC[park][kind] is not None)
            w.setEnabled(has and self.data is not None)
            if self.data is not None and park != "JOINT":
                w.blockSignals(True)
                w.setValue(float(self.data.caps[kind][park]))
                w.blockSignals(False)
            elif park == "JOINT":
                w.blockSignals(True)
                w.setValue(0.0)
                w.blockSignals(False)

    def _on_cap_changed(self) -> None:
        """修改装机后立即写入数据容器（pu 不变，kW = pu × 装机）。"""
        if self.data is None:
            return
        park = self._cur_park()
        if park == "JOINT":
            return
        if PU_SRC[park]["wind"] is not None:
            self.data.caps["wind"][park] = self.sp_windcap.value()
        if PU_SRC[park]["pv"] is not None:
            self.data.caps["pv"][park] = self.sp_pvcap.value()
        _, W, V = self.data.arrays(park, self._cur_month())
        gen_max = float((W + V).max()) if len(W) else 0.0
        self.status.showMessage(
            f"已更新园区{park} 装机：风电 {self.data.caps['wind'][park]:.0f} kW / "
            f"光伏 {self.data.caps['pv'][park]:.0f} kW（出力峰值 {gen_max:.0f} kW）")

    def _after_data_change(self, msg: str) -> None:
        self.lbl_src.setText("当前：" + self.data.source)
        self._sync_park()
        self._log(msg)
        for c in self.data.checks:
            self._log("  " + c)
        self.payload = {"view": "data", "tables": {"数据检查": list(self.data.src_checks)},
                        "status": msg}
        self.view = "data"
        self._fill(self.tables["数据检查"], self.data.src_checks)
        self.tabs.setCurrentWidget(self.tables["数据检查"])
        self._render()
        self.status.showMessage(msg)

    def on_load_excel(self) -> None:
        """载入用户自备的 Excel：按文件名自动识别附件1/2/3，未识别的按选择顺序补齐。"""
        start = _osp.dirname(self.data.src_checks and "") or os.getcwd()
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择 Excel 数据文件（可多选：附件1 / 附件2 / 附件3）", start,
            "Excel 文件 (*.xlsx *.xls *.xlxs)")
        if not paths:
            return
        b = [_osp.basename(p) for p in paths]
        p1 = next((p for p in paths if "附件1" in _osp.basename(p)), None)
        p2 = next((p for p in paths if ("附件2" in _osp.basename(p)
                                        or _osp.basename(p).lower().startswith("att2"))), None)
        p3 = next((p for p in paths if "附件3" in _osp.basename(p)), None)
        ordered = [x for x in (p1, p2, p3) if x]
        for p in paths:
            if p not in ordered:
                ordered.append(p)
        try:
            self.data = excel_data(ordered)
        except Exception as e:
            QMessageBox.critical(self, "载入失败", f"{type(e).__name__}: {e}\n\n所选：{b}")
            return
        self._after_data_change("已载入 Excel 数据：" + "、".join(b))

    def on_use_embedded(self) -> None:
        """恢复使用代码内置的数据模板。"""
        try:
            self.data = embedded_data()
        except Exception as e:
            QMessageBox.critical(self, "恢复失败", f"{type(e).__name__}: {e}")
            return
        self._after_data_change("已恢复内置数据模板。")

    def _need_data(self) -> bool:
        if self.data is None:
            QMessageBox.information(self, "提示", "请先点击「① 数据检查」读取数据。")
            return False
        return True

    # ------------------------------------------------------------------ 后台任务
    def _run_async(self, fn, title: str) -> None:
        if self.worker is not None and self.worker.isRunning():
            QMessageBox.information(self, "提示", "已有任务在运行，请先等待完成或点「停止计算」。")
            return
        clear_cancel()
        self._busy(True, title)
        self._last_pct = -1
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.worker = Worker(fn)
        self.worker.log.connect(self._log)
        self.worker.prog.connect(self._on_progress)
        self.worker.done.connect(self._on_done)
        self.worker.fail.connect(self._on_fail)
        self.worker.stopped.connect(self._on_cancelled)
        self.worker.start()

    def _busy(self, on: bool, which: str = "") -> None:
        for b in self.buttons:
            b.setEnabled(not on)
        self.btn_stop.setEnabled(on)
        if not on:
            self.progress.setVisible(False)
        self.status.showMessage(f"正在{which}…（可点「停止计算」中止）" if on else "就绪")

    def _on_progress(self, done: int, total: int, msg: str) -> None:
        total = max(1, total)
        pct = max(0, min(100, int(100 * done / total)))
        if pct == self._last_pct:
            return
        self._last_pct = pct
        self.progress.setValue(pct)
        self.status.showMessage(f"正在计算 {msg}… {pct}%")

    def _on_cancelled(self) -> None:
        self._busy(False)
        self._log("计算已中止（本次结果未更新）。可放宽网格或改用「粗」精度后重试。")
        self.status.showMessage("已中止")

    def on_stop(self) -> None:
        """中止当前长任务（网格搜索在每个网格点检查一次中止标志）。"""
        if self.worker is not None and self.worker.isRunning():
            request_cancel()
            self.btn_stop.setEnabled(False)
            self._log("已发出中止请求，正在收尾…")
        else:
            self.status.showMessage("当前没有正在运行的计算")

    def _log(self, s: str) -> None:
        ts = _dt.datetime.now().strftime("%H:%M:%S")
        self.tab_log.appendPlainText(f"[{ts}] {s}")
        self.tab_log.moveCursor(QTextCursor.End)

    def _on_fail(self, msg: str) -> None:
        self._busy(False)
        self._log("执行失败：" + msg.splitlines()[0])
        QMessageBox.critical(self, "执行失败", msg)

    def _on_done(self, payload: Dict[str, Any]) -> None:
        self._busy(False)
        self.payload = payload
        for name, rows in (payload.get("tables") or {}).items():
            if name in self.tables:
                self._fill(self.tables[name], rows)
        self.view = payload.get("view", self.view)
        if payload.get("focus") in self.tables:
            self.tabs.setCurrentWidget(self.tables[payload["focus"]])
        try:                     # 绘图异常不得中断主流程（PyQt 槽内未捕获异常会终止进程）
            self._render()
        except Exception as e:
            self._log(f"绘图失败（结果已计算）：{type(e).__name__}: {e}")
        if SOLVER["rule_fallback"]:
            self._log("⚠ 本次计算使用了纯 Python 规则策略（本机无可用 LP 求解器）。"
                      "在「单一电价 + 不允许售电」设定下该结果即精确最优；"
                      "若启用了分时电价或允许售电，结果为近似值。")
            self._apply_solver_capability()
        self.status.showMessage(payload.get("status", "完成"))

    # ------------------------------------------------------------------ 表格填充
    @staticmethod
    def _fill(table: QTableWidget, rows: Sequence[Dict[str, Any]]) -> None:
        table.clear()
        if not rows:
            table.setRowCount(0)
            table.setColumnCount(0)
            return
        keys: List[str] = []
        for r in rows:
            for k in r:
                if k not in keys:
                    keys.append(k)
        table.setColumnCount(len(keys))
        table.setRowCount(len(rows))
        table.setHorizontalHeaderLabels(keys)
        for i, r in enumerate(rows):
            for j, k in enumerate(keys):
                val = r.get(k, "")
                if isinstance(val, float):
                    val = f"{val:.6g}"
                item = QTableWidgetItem(str(val))
                if k in ("结论", "是否最优", "备注"):
                    s = str(val)
                    if ("未通过" in s) or ("失败" in s) or s == "否":
                        item.setForeground(QColor("#c62828"))
                        f = item.font(); f.setBold(True); item.setFont(f)
                    elif ("通过" in s) or s == "是":
                        item.setForeground(QColor("#1a7f37"))
                    elif "提示" in s:
                        item.setForeground(QColor("#e08b00"))
                table.setItem(i, j, item)
        table.resizeColumnsToContents()
        for j in range(table.columnCount()):
            if table.columnWidth(j) > 240:
                table.setColumnWidth(j, 240)

    # ------------------------------------------------------------------ 绘图
    def _style(self, ax, title, ylab, xlab="时刻 (h)") -> None:
        ax.set_title(title, fontsize=10.5, pad=6)
        ax.set_xlabel(xlab, fontsize=9)
        ax.set_ylabel(ylab, fontsize=9)
        if xlab.startswith("时刻"):
            ax.set_xticks(np.arange(0, 24, 2))
            ax.set_xlim(-0.4, 23.4)
        ax.grid(alpha=0.3, lw=0.6, ls=":")
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)

    def _render(self) -> None:
        if not HAS_MPL or self.fig is None:
            return
        self.fig.clear()
        getattr(self, f"_draw_{self.view}", self._draw_welcome)()
        try:
            self.fig.tight_layout()
        except Exception:
            pass
        self.canvas.draw()

    def _draw_welcome(self) -> None:
        ax = self.fig.add_subplot(111)
        ax.axis("off")
        ax.text(0.5, 0.58, f"{APP_NAME}", ha="center", va="center", fontsize=17, weight="bold")
        ax.text(0.5, 0.47, APP_VERSION, ha="center", va="center", fontsize=12, color="#666")
        ax.text(0.5, 0.33,
                "按顶部 ①→⑦ 顺序运行：数据检查 → 无储能基线 → 储能运行模拟 →\n"
                "最优运行策略 → 容量优化 → 联合运营 → 敏感性分析；⑧ 结果自检可随时校验。",
                ha="center", va="center", fontsize=10.5, color="#444", linespacing=1.8)

    def _draw_data(self) -> None:
        d, m = self.data, self._cur_month()
        tag = "附件2 典型日" if m <= 0 else f"附件3 {m} 月"
        axes = self.fig.subplots(2, 2)
        for p in PARKS:
            L, W, V = d.arrays(p, m)
            axes[0][0].plot(X24, L, "-o", ms=3, lw=1.7, color=PCOLOR[p], label=f"园区{p}")
            axes[0][1].plot(X24, W, "-", lw=1.6, color=PCOLOR[p], label=f"园区{p} 风电")
            axes[0][1].plot(X24, V, "--", lw=1.6, color=PCOLOR[p], alpha=.8, label=f"园区{p} 光伏")
            axes[1][0].plot(X24, L - W - V, "-s", ms=3, lw=1.6, color=PCOLOR[p], label=f"园区{p} 净负荷")
        axes[1][0].axhline(0, color="#666", lw=0.9, ls="--")
        self._style(axes[0][0], "典型日负荷曲线", "负荷 (kW)")
        self._style(axes[0][1], "风光出力曲线", "出力 (kW)")
        self._style(axes[1][0], "净负荷曲线（负荷−风光，负值=弃电）", "净负荷 (kW)")
        for a in (axes[0][0], axes[0][1], axes[1][0]):
            a.legend(frameon=False, fontsize=8, ncol=2)

        ax = axes[1][1]
        if m <= 0:
            z = (np.asarray(d.m_pu["A_pv"], float) * d.caps["pv"]["A"]).T
            im = ax.imshow(z, aspect="auto", cmap="viridis", origin="lower")
            ax.set_yticks(np.arange(0, 24, 3)); ax.set_ylabel("时刻 (h)", fontsize=9)
            ax.set_xticks(np.arange(0, 12)); ax.set_xticklabels([f"{i+1}" for i in range(12)], fontsize=8)
            ax.set_title("园区A 光伏 12 个月出力热力图 (kW)", fontsize=10.5)
            self.fig.colorbar(im, ax=ax, shrink=.9)
        else:
            for p in PARKS:
                L, W, V = d.arrays(p, m)
                ax.bar(X24 - .25 + .25 * PARKS.index(p), L - W - V, .25,
                       color=PCOLOR[p], label=f"园区{p} 净负荷")
            ax.axhline(0, color="#666", lw=.9)
            self._style(ax, f"{m} 月三园区净负荷", "净负荷 (kW)")
            ax.legend(frameon=False, fontsize=8)
        self.fig.suptitle(f"数据检查与典型日曲线（{tag}）", fontsize=12.5)

    def _draw_baseline(self) -> None:
        rows = self.payload.get("tables", {}).get("无储能基线", [])
        if not rows:
            self._draw_welcome(); return
        by = {r["指标"]: r for r in rows}
        cols = [k for k in rows[0].keys() if k != "指标"]       # 形如「园区A」「园区B」…
        axes = self.fig.subplots(1, 2)
        labels = cols
        keys = [("购电量", "#C0392B"), ("绿电自用电量", "#27AE60"), ("弃电量", "#E67E22")]
        bottom = np.zeros(len(cols))
        for k, c in keys:
            vals = np.array([float(by[k][p]) for p in cols])
            axes[0].bar(labels, vals, .55, bottom=bottom, color=c, label=k)
            bottom += vals
        axes[0].set_ylabel("电量 (kWh/日)", fontsize=9)
        axes[0].set_title("无储能电量构成", fontsize=10.5)
        axes[0].legend(frameon=False, fontsize=8.5)
        axes[0].grid(alpha=.3, lw=.6, ls=":", axis="y")

        uc = [float(by["单位电量平均供电成本"][p]) for p in cols]
        bb = axes[1].bar(labels, uc, .55, color=["#C0392B", "#1F6FB4", "#1E8449"])
        axes[1].bar_label(bb, fmt="%.4f", fontsize=9, padding=2)
        axes[1].axhline(1.0, color="#666", ls="--", lw=1, label="直接购电价 1.0")
        axes[1].set_ylabel("单位电量平均供电成本 (元/kWh)", fontsize=9)
        axes[1].set_ylim(0, max(uc) * 1.2)
        axes[1].set_title("单位电量平均供电成本", fontsize=10.5)
        axes[1].legend(frameon=False, fontsize=8.5)
        axes[1].grid(alpha=.3, lw=.6, ls=":", axis="y")
        self.fig.suptitle("无储能基线（确定性算术）", fontsize=12.5)

    def _draw_compare(self) -> None:
        tab = self.payload.get("cmp", {})
        if not tab:
            self._draw_welcome(); return
        panels = [("总供电成本", "总供电成本 (元/日)"), ("单位电量平均供电成本", "单位成本 (元/kWh)"),
                  ("购电量", "购电量 (kWh/日)"), ("弃电量", "弃电量 (kWh/日)")]
        tags = ["无储能", "规则策略", "LP最优"]
        cols3 = ["#95A5A6", "#E67E22", "#1F6FB4"]
        parks = [p for p in PARKS if p in tab]
        axes = self.fig.subplots(2, 2)
        w = 0.8 / len(tags)
        xs = np.arange(len(parks))
        for k, (col, ylab) in enumerate(panels):
            ax = axes[k // 2][k % 2]
            for i, (tag, cc) in enumerate(zip(tags, cols3)):
                vals = [tab[p][tag][col] for p in parks]
                bb = ax.bar(xs + (i - 1) * w, vals, w, color=cc, label=tag)
                ax.bar_label(bb, fmt="%.4f" if "单位" in col else "%.0f", fontsize=7.5, padding=1.5)
            ax.set_xticks(xs); ax.set_xticklabels(["园区" + p for p in parks])
            ax.set_ylabel(ylab, fontsize=9); ax.set_title(ylab, fontsize=10.5)
            ax.grid(alpha=.3, lw=.6, ls=":", axis="y")
            for s in ("top", "right"):
                ax.spines[s].set_visible(False)
            yl = ax.get_ylim(); ax.set_ylim(yl[0], yl[1] * 1.15)
        axes[0][0].legend(frameon=False, fontsize=8.5, ncol=3)
        self.fig.suptitle(f"三方案对比（储能 {fnum(self.payload.get('P', 0), 0)} kW / "
                          f"{fnum(self.payload.get('E', 0), 0)} kWh）", fontsize=12.5)

    def _draw_dispatch(self) -> None:
        res = self.payload.get("dispatch")
        if not res:
            self._draw_welcome(); return
        L = res["L"]
        r = res["res"]
        axes = self.fig.subplots(2, 1, sharex=True)
        ax = axes[0]
        su = r.wu + r.vu
        ax.bar(X24, su, .72, color="#27AE60", label="绿电直供")
        ax.bar(X24, r.buy, .72, bottom=su, color="#C0392B", label="电网购电")
        ax.bar(X24, -r.ch, .72, color="#2E86C1", label="储能充电")
        ax.bar(X24, r.dis, .72, color="#8E44AD", label="储能放电")
        if float(np.sum(r.ws + r.ps)) > 1e-6:
            ax.bar(X24, r.ws + r.ps, .72, color="#16A085", label="外送售电")
        ax.plot(X24, L, "k--", lw=1.5, label="负荷")
        self._style(ax, f"逐时功率（{r.label}·{res['park']}·{res['tag']}）", "功率 (kW)")
        ax.legend(frameon=False, fontsize=8.5, ncol=3, loc="best")
        ax2 = axes[1]
        ax2.plot(X24, r.soc, "-o", ms=4, lw=2, color="#8E44AD")
        if r.E > 0:
            ax2.axhline(self._spec().soc_min * r.E, color="#C0392B", ls=":", lw=1.2,
                        label=f"SOC 下限 {self._spec().soc_min * 100:.0f}%（{self._spec().soc_min * r.E:.0f} kWh）")
            ax2.axhline(self._spec().soc_max * r.E, color="#1F6FB4", ls=":", lw=1.2,
                        label=f"SOC 上限 {self._spec().soc_max * 100:.0f}%（{self._spec().soc_max * r.E:.0f} kWh）")
            ax2.fill_between(X24, self._spec().soc_min * r.E, self._spec().soc_max * r.E,
                             color="#8E44AD", alpha=.06)
        self._style(ax2, "储能荷电状态 SOC", "储能存量 (kWh)")
        ax2.legend(frameon=False, fontsize=8.5, loc="upper left")
        self.fig.suptitle("运行策略与储能轨迹", fontsize=12.5)

    def _draw_capacity(self) -> None:
        caps = self.payload.get("caps")
        if not caps:
            self._draw_welcome(); return
        parks = self.payload.get("parks") or list(caps.keys())
        n = len(parks)
        axes = self.fig.subplots(2, n, squeeze=False)
        for j, p in enumerate(parks):
            r = caps[p]
            b = r["best"]
            Pl, El, S = r["P_list"], r["E_list"], r["surf"]
            Pc, Ec = np.meshgrid(Pl, El, indexing="ij")

            ax = axes[0][j]
            cs = ax.contourf(Pc, Ec, S, levels=22, cmap="viridis_r")
            c2 = ax.contour(Pc, Ec, S, levels=8, colors="w", linewidths=.5, alpha=.55)
            ax.clabel(c2, fmt="%.0f", fontsize=6)
            if b["P"] > 0:
                ax.plot(b["P"], b["E"], "*", ms=15, color="#E74C3C", mec="w", mew=1.1,
                        label=f"最优 ({b['P']:.1f},{b['E']:.1f})")
            else:
                ax.plot(0, 0, "*", ms=15, color="#E74C3C", mec="w", mew=1.1, label="最优：不配储能")
            ax.plot(50, 100, "P", ms=10, color="#F1C40F", mec="k", mew=.8, label="题目 (50,100)")
            ax.set_xlabel("储能功率 P (kW)", fontsize=8.5)
            ax.set_ylabel("储能容量 E (kWh)", fontsize=8.5)
            ax.set_title(f"园区{p} 年化成本曲面", fontsize=10)
            ax.tick_params(labelsize=8)
            ax.legend(frameon=False, fontsize=7, loc="upper right")
            self.fig.colorbar(cs, ax=ax, shrink=.85).ax.tick_params(labelsize=6.5)

            # 成本分解：沿 P 轴取每档功率下的最优容量，反推投资/运行/合计
            ax = axes[1][j]
            xsP, inv, op, tot = [], [], [], []
            for i, P in enumerate(Pl):
                col = S[i, :]
                if np.all(np.isnan(col)):
                    continue
                k = int(np.nanargmin(col))
                E = El[k]
                tv = float(col[k])
                iv = 0.0 if P <= 0 else (r["spec"].cost_p * P + r["spec"].cost_e * E) / r["spec"].life_y
                xsP.append(P); inv.append(iv); op.append(tv - iv); tot.append(tv)
            ax.plot(xsP, tot, "-o", ms=3.5, lw=1.8, color="#1F6FB4", label="年化总成本")
            ax.plot(xsP, op, "--s", ms=3.5, lw=1.4, color="#27AE60", label="年运行成本")
            ax.plot(xsP, inv, ":^", ms=3.5, lw=1.4, color="#E67E22", label="年折旧投资")
            if b["P"] > 0:
                ax.axvline(b["P"], color="#E74C3C", lw=1.3, ls="-.",
                           label=f"最优 P={b['P']:.0f} kW")
            else:
                ax.annotate("最优：不配储能", xy=(xsP[1] if len(xsP) > 1 else 20,
                            tot[1] if len(tot) > 1 else 0), xytext=(.28, .70),
                            textcoords="axes fraction", fontsize=8, color="#E74C3C",
                            arrowprops=dict(arrowstyle="->", color="#E74C3C",
                                            connectionstyle="arc3,rad=.25"))
            ax.set_xlabel("储能功率 P (kW)", fontsize=8.5)
            ax.set_ylabel("成本 (元/年)", fontsize=8.5)
            ax.set_title(f"园区{p} 成本分解", fontsize=10)
            ax.grid(alpha=.3, lw=.6, ls=":")
            ax.tick_params(labelsize=8)
            for s in ("top", "right"):
                ax.spines[s].set_visible(False)
            ax.legend(frameon=False, fontsize=7)
        self.fig.suptitle("储能功率/容量优化（目标：年化总成本最小）", fontsize=12.5)

    def _draw_joint(self) -> None:
        rows = self.payload.get("tables", {}).get("联合运营", [])
        if not rows:
            self._draw_welcome(); return
        labels = [r["情景"].replace("运营 · ", "\n") for r in rows]
        vals = [float(r["年化总成本"]) for r in rows]
        cc = ["#95A5A6", "#1F6FB4", "#7FB3D5", "#95A5A6", "#E67E22", "#C0392B"]
        axes = self.fig.subplots(1, 2)
        bb = axes[0].bar(range(len(vals)), vals, .6, color=cc[:len(vals)])
        axes[0].bar_label(bb, fmt="%.0f", fontsize=8, padding=2)
        axes[0].set_xticks(range(len(vals)))
        axes[0].set_xticklabels(labels, fontsize=7.5)
        axes[0].set_ylabel("年化总成本 (元/年)", fontsize=9)
        axes[0].set_title("六情景年化总成本对比", fontsize=10.5)
        axes[0].grid(alpha=.3, lw=.6, ls=":", axis="y")
        axes[0].set_ylim(min(vals) * .96, max(vals) * 1.04)
        for s in ("top", "right"):
            axes[0].spines[s].set_visible(False)

        m = self._cur_month()
        ci = np.zeros(24); cj = np.zeros(24)
        for p in PARKS:
            L, W, V = self.data.arrays(p, m)
            ci += np.maximum(W + V - L, 0)
        L, W, V = self.data.joint_arrays(m)
        cj += np.maximum(W + V - L, 0)
        axes[1].bar(X24 - .2, ci, .38, color="#C0392B", label="独立运营合计")
        axes[1].bar(X24 + .2, cj, .38, color="#1E8449", label="联合运营")
        self._style(axes[1], f"弃电功率对比（联合消纳 {ci.sum() - cj.sum():.0f} kWh/日）", "弃电功率 (kW)")
        axes[1].legend(frameon=False, fontsize=9)
        self.fig.suptitle("三园区联合运营经济性与机制", fontsize=12.5)

    def _draw_sens(self) -> None:
        rows = self.payload.get("tables", {}).get("敏感性", [])
        if not rows:
            self._draw_welcome(); return
        df = rows
        params = []
        for r in df:
            if r["参数"] not in params:
                params.append(r["参数"])
        axes = self.fig.subplots(2, 2)
        for k, pname in enumerate(params[:4]):
            ax = axes[k // 2][k % 2]
            labs = []
            for r in df:
                if r["参数"] == pname and r["取值"] not in labs:
                    labs.append(r["取值"])
            xs = np.arange(len(labs))
            axb = ax.twinx()
            for p in PARKS:
                ys = []
                for lab in labs:
                    v = next((r for r in df if r["参数"] == pname and r["取值"] == lab and r["园区"] == p), None)
                    ys.append(float(v["年化总成本"]) / 1e4 if v else np.nan)
                ax.plot(xs, ys, "-o", ms=4, lw=1.8, color=PCOLOR[p], label=f"园区{p}")
                ysE = []
                for lab in labs:
                    v = next((r for r in df if r["参数"] == pname and r["取值"] == lab and r["园区"] == p), None)
                    ysE.append(float(v["最优E"]) if v else np.nan)
                axb.plot(xs, ysE, ":", lw=1.1, color=PCOLOR[p], alpha=.75)
            ax.set_xticks(xs); ax.set_xticklabels(labs, fontsize=8)
            ax.set_xlabel(pname, fontsize=9)
            ax.set_ylabel("年化总成本 (万元/年)", fontsize=9)
            axb.set_ylabel("最优 E (kWh)", fontsize=8)
            axb.tick_params(labelsize=7.5)
            ax.set_title(pname, fontsize=10.5)
            ax.grid(alpha=.3, lw=.6, ls=":")
            ax.spines["top"].set_visible(False); axb.spines["top"].set_visible(False)
            if k == 0:
                ax.legend(frameon=False, fontsize=8, loc="upper left")
        self.fig.suptitle("敏感性分析：参数扰动下的最优配置与年化总成本（实线=成本，点线=最优E）",
                           fontsize=11.5)

    def _draw_validate(self) -> None:
        rows = self.payload.get("tables", {}).get("结果自检", [])
        if not rows:
            self._draw_welcome(); return
        ax = self.fig.add_subplot(111)
        ax.axis("off")
        # 兼容两种自检行结构：单项校验(实测/容差) 与 回归自检(说明)
        # 把汇总行并入同一文本块，由 matplotlib 统一排版，避免行数与字号变化时发生重叠
        n = len(rows)
        fs = 11.0 if n <= 8 else (10.0 if n <= 12 else 9.0)
        lines = [
            f"{'√' if r['结论'] == '通过' else '×'}  {r['检查项']}："
            f"{r.get('实测', r.get('说明', ''))}"
            + (f"（容差 {r['容差']}）" if r.get("容差") else "")
            for r in rows]
        npass = sum(1 for r in rows if r["结论"] == "通过")
        lines += ["", f"合计 {npass}/{n} 项通过"
                  + ("" if npass == n else "（存在未通过项，请检查参数与数据）")]
        ax.text(0.01, 0.98, "结果自检 / 一键回归自检", fontsize=13, weight="bold", va="top")
        ax.text(0.01, 0.86, "\n".join(lines), fontsize=fs, va="top",
                linespacing=1.55, family="Microsoft YaHei")

    # ------------------------------------------------------------------ 槽函数
    def _auto_load(self) -> None:
        """默认装载代码内置的数据模板（无需任何外部文件）。"""
        try:
            self.data = embedded_data()
        except Exception as e:
            self._log("内置数据初始化失败：" + str(e))
            QMessageBox.warning(self, "数据初始化失败", str(e))
            return
        self.lbl_src.setText("当前：" + self.data.source)
        self._sync_park()
        self._log(f"数据就绪：{self.data.source}，{len(self.data.hours)} 时段，"
                  f"三园区负荷与风光出力按装机折算为 kW（共 {len(self.data.checks)} 项检查）。")
        self._log("提示：可用「载入 Excel 数据…」换成自己的附件；也可直接改左侧"
                  "「风电/光伏装机」「负荷缩放」，再重跑 ①~⑦ 对比结果。")

    def on_data(self) -> None:
        if self.data is None:
            self._auto_load()
        if self.data is None:
            return
        rows = list(self.data.src_checks)
        self.view = "data"
        self.payload = {"view": "data", "tables": {"数据检查": rows},
                        "status": f"数据检查完成（{len(rows)} 项）"}
        self._fill(self.tables["数据检查"], rows)
        self.tabs.setCurrentWidget(self.tables["数据检查"])
        self._render()
        self._log("① 数据检查完成，检查明细见「数据检查」页签。")

    def on_baseline(self) -> None:
        if not self._need_data():
            return
        sp = self._spec(); m = self._cur_month()
        parks = PARKS                      # 3.3-4 要求三园区汇总成表，故始终计算三个园区
        rows = []
        for p in parks:
            L, W, V = self._arrays(p, m)
            r = baseline(L, W, V, sp, p)
            mt = r.metrics(L, W, V, sp)
            rows.append({"指标": "园区" + p,
                         **{k: round(v, 4) if isinstance(v, float) else v
                            for k, v in mt.items()}})
        # 转置为「指标 × 园区」便于阅读
        keys = [k for k in rows[0] if k != "指标"]
        table = [{"指标": k, **{r["指标"]: r[k] for r in rows}} for k in keys]
        self.view = "baseline"
        self.payload = {"view": "baseline", "tables": {"无储能基线": table},
                        "status": "无储能基线完成"}
        self._fill(self.tables["无储能基线"], table)
        self.tabs.setCurrentWidget(self.tables["无储能基线"])
        self._render()
        for r in rows:
            self._log(f"② 无储能基线 {r['指标']}：单位成本 "
                      f"{fnum(r['单位电量平均供电成本'], 4)} 元/kWh，弃电率 {100 * r['弃电率']:.2f}%")

    def on_simulate(self) -> None:
        if not self._need_data():
            return
        sp = self._spec(); m = self._cur_month()
        P, E = self.sp_P.value(), self.sp_E.value()
        parks = PARKS                      # 3.3-5 要求各园区分别配置储能，故始终覆盖三个园区
        cmp_tab: Dict[str, Any] = {}
        rows = []
        last = None
        for p in parks:
            L, W, V = self._arrays(p, m)
            r0 = baseline(L, W, V, sp, p)
            r1 = rule_strategy(L, W, V, P, E, sp, p) if (P > 0 and E > 0) else None
            r2 = solve_operation(L, W, V, P, E, sp, p, milp=self.chk_milp.isChecked(),
                                 time_limit=self.sp_tl.value()) if (P > 0 and E > 0) else r0
            cmp_tab[p] = {}
            for tag, r in (("无储能", r0), ("规则策略", r1 or r0), ("LP最优", r2)):
                mt = r.metrics(L, W, V, sp)
                cmp_tab[p][tag] = mt
                rows.append({"园区": p, "方案": tag, "求解器": r.engine,
                             "购电量": round(mt["购电量"], 2), "弃电量": round(mt["弃电量"], 2),
                             "充电量": round(mt["充电量"], 2), "放电量": round(mt["放电量"], 2),
                             "总供电成本": round(mt["总供电成本"], 2),
                             "单位电量平均供电成本": round(mt["单位电量平均供电成本"], 4),
                             "新能源渗透率": f"{100 * mt['新能源渗透率']:.2f}%",
                             "弃电率": f"{100 * mt['弃电率']:.2f}%"})
            last = (p, L, W, V, r2)
            b, rl, lp = (cmp_tab[p][t]["总供电成本"] for t in ("无储能", "规则策略", "LP最优"))
            self._log(f"③ 园区{p}（{P:.0f}kW/{E:.0f}kWh）：无储能 {b:.2f} → 规则 {rl:.2f}"
                      f"（↓{b - rl:.2f}）→ LP {lp:.2f}（再↓{rl - lp:.4f}）元/日")
        self.view = "compare"
        self.payload = {"view": "compare", "cmp": cmp_tab, "P": P, "E": E,
                        "tables": {"方案对比": rows}, "status": "储能运行模拟完成"}
        self._fill(self.tables["方案对比"], rows)
        self.tabs.setCurrentWidget(self.tables["方案对比"])
        self._render()
        if last is not None:
            sel = self._cur_park()
            self._store_dispatch(sel if sel in PARKS else PARKS[0], m, sp, P, E)

    def _store_dispatch(self, park, m, sp, P, E):
        L, W, V = self._arrays(park, m)
        r = solve_operation(L, W, V, P, E, sp, park, milp=self.chk_milp.isChecked(),
                            time_limit=self.sp_tl.value())
        self.payload["dispatch"] = {"park": park, "tag": self.cb_src.currentText(),
                                    "L": L, "W": W, "V": V, "res": r}
        rows = []
        for t in range(24):
            rows.append({"时刻": f"{t:02d}:00", "负荷(kW)": round(L[t], 2),
                         "风电(kW)": round(W[t], 2), "光伏(kW)": round(V[t], 2),
                         "风电入系统": round(r.wu[t], 2), "光伏入系统": round(r.vu[t], 2),
                         "购电": round(r.buy[t], 2), "充电": round(r.ch[t], 2),
                         "放电": round(r.dis[t], 2), "SOC(kWh)": round(r.soc[t], 2)})
        self._fill(self.tables["调度明细"], rows)

    def on_optimal(self) -> None:
        """④ 最优运行策略：对给定 (P,E) 求 LP 最优调度并展示 SOC 轨迹。"""
        if not self._need_data():
            return
        sp = self._spec(); m = self._cur_month(); park = self._cur_park()
        P, E = self.sp_P.value(), self.sp_E.value()
        if park == "JOINT":
            L, W, V = self.data.joint_arrays(m)
            r = solve_operation(L, W, V, P, E, sp, "JOINT", milp=self.chk_milp.isChecked())
        else:
            L, W, V = self._arrays(park, m)
            r = solve_operation(L, W, V, P, E, sp, park, milp=self.chk_milp.isChecked())
        mt = r.metrics(L, W, V, sp)
        rows = [{"时刻": f"{t:02d}:00", "负荷(kW)": round(L[t], 2), "风电(kW)": round(W[t], 2),
                 "光伏(kW)": round(V[t], 2), "购电": round(r.buy[t], 2),
                 "充电": round(r.ch[t], 2), "放电": round(r.dis[t], 2),
                 "SOC(kWh)": round(r.soc[t], 2)} for t in range(24)]
        self.view = "dispatch"
        self.payload = {"view": "dispatch",
                        "dispatch": {"park": park, "tag": self.cb_src.currentText(),
                                     "L": L, "W": W, "V": V, "res": r},
                        "tables": {"调度明细": rows}, "status": "最优运行策略完成"}
        self._fill(self.tables["调度明细"], rows)
        self.tabs.setCurrentWidget(self.tables["调度明细"])
        self._render()
        self._log(f"④ {park} 最优运行策略（储能 {P:.0f}kW/{E:.0f}kWh，{r.engine}）："
                  f"日成本 {mt['总供电成本']:.2f} 元，单位成本 {mt['单位电量平均供电成本']:.4f} 元/kWh，"
                  f"弃电率 {100 * mt['弃电率']:.2f}%，{r.status}")

    def on_capacity(self) -> None:
        if not self._need_data():
            return
        sp = self._spec(); m = self._cur_month(); park = self._cur_park()
        Pmax, pstep = self.sp_pmax.value(), max(1.0, self.sp_pstep.value())
        Emax, estep = self.sp_emax.value(), max(1.0, self.sp_estep.value())
        refine = self.sp_refine.value()
        Pl, El = build_grid(pstep, estep, Pmax, Emax)
        # 3.4-2 要求逐园区优化；仅当显式选择「三园区联合」时才优化联合系统
        parks = ("JOINT",) if park == "JOINT" else PARKS
        budget = solve_budget_note(estimate_solves(Pl, El, refine), len(parks))

        def job(log, prog):
            log(f"⑤ 容量优化：{len(parks)} 个分析对象 × {len(Pl)}×{len(El)} 网格，{budget}")
            rows, caps = [], {}
            for idx, p in enumerate(parks):
                L, W, V = self._arrays(p, m)
                log(f"  对象{p} 开始（第 {idx + 1}/{len(parks)} 个）…")

                def sub(d, t, pk, _i=idx, _p=p):
                    prog(int((_i + d / max(1, t)) * 1000), len(parks) * 1000, f"对象{_p}")

                r = optimize_capacity(L, W, V, sp, park=p, P_list=Pl, E_list=El,
                                      n_refine=refine, log=log, progress=sub, verify=True)
                caps[p] = r
                b = r["best"]
                c50 = op_cost(L, W, V, 50, 100, r["spec"]) if (50 <= Pmax and 100 <= Emax) else None
                inv50 = sp.cost_p * 50 + sp.cost_e * 100
                a50 = (c50 * sp.days + inv50 / sp.life_y) if c50 is not None else np.nan
                a0 = r["base_cost"] * sp.days
                rows.append({"园区": p, "最优P(kW)": round(b["P"], 2), "最优E(kWh)": round(b["E"], 2),
                             "E/P时长(h)": round(b["E"] / b["P"], 2) if b["P"] else "-",
                             "最优年化总成本": round(b["年化总成本"], 1),
                             "无储能年化": round(a0, 1),
                             "题目50/100年化": round(a50, 1) if a50 == a50 else "-",
                             "最优较50/100省": round(a50 - b["年化总成本"], 1) if a50 == a50 else "-",
                             "最优较无储能省": round(a0 - b["年化总成本"], 1),
                             "是否最优=不配储能": "是" if b["P"] == 0 else "否",
                             "求解器偏差(元)": fnum(b.get("求解器偏差", 0), 6)})
                extra50 = f"，50/100 方案 {a50:.1f}" if a50 == a50 else ""
                log(f"  对象{p}：最优 P={b['P']:.2f} kW / E={b['E']:.2f} kWh，"
                    f"年化 {b['年化总成本']:.1f} 元；无储能 {a0:.1f}{extra50}")
            return {"view": "capacity", "caps": caps, "parks": list(parks),
                    "tables": {"容量优化": rows},
                    "status": "容量优化完成"}

        self._run_async(job, "容量优化")

    def on_joint(self) -> None:
        if not self._need_data():
            return
        sp = self._spec(); m = self._cur_month()

        Pl, El = build_grid(max(10.0, self.sp_pstep.value()), max(25.0, self.sp_estep.value()),
                            min(260.0, self.sp_pmax.value()), min(550.0, self.sp_emax.value()))
        units = len(PARKS) + 1

        def job(log, prog):
            rows = []
            check_cancel()
            # 独立：无储能
            ind0 = sum(op_cost(*self._arrays(p, m), 0, 0, sp) for p in PARKS)
            rows.append({"情景": "独立运营 · 无储能", "储能配置": "各园区均无",
                         "日运行成本": round(ind0, 2), "年运行成本": round(ind0 * sp.days, 1),
                         "年折旧投资": 0.0, "年化总成本": round(ind0 * sp.days, 1)})
            # 独立：各配 50/100
            ind50 = sum(op_cost(*self._arrays(p, m), 50, 100, sp) for p in PARKS)
            inv50 = (sp.cost_p * 50 + sp.cost_e * 100) * 3 / sp.life_y
            rows.append({"情景": "独立运营 · 各配 50/100", "储能配置": "三园区各 50 kW/100 kWh",
                         "日运行成本": round(ind50, 2), "年运行成本": round(ind50 * sp.days, 1),
                         "年折旧投资": round(inv50, 1), "年化总成本": round(ind50 * sp.days + inv50, 1)})
            log(f"⑥ 联合运营：{units} 个容量优化 × {len(Pl)}×{len(El)} 网格，"
                f"{solve_budget_note(estimate_solves(Pl, El, 1), units)}")
            # 独立：各配最优（批量调用，关闭 PuLP 复算以提速）
            ind_best, ind_inv, cfg = 0.0, 0.0, []
            for k, p in enumerate(PARKS):
                def sub(d, t, pk, _k=k, _p=p):
                    prog(int((_k + d / max(1, t)) * 1000), units * 1000, f"独立·园区{_p}")
                r = optimize_capacity(*self._arrays(p, m), sp=sp, park=p, P_list=Pl, E_list=El,
                                      n_refine=1, verify=False, progress=sub)
                b = r["best"]
                ind_best += b["日运行成本"]
                ind_inv += (sp.cost_p * b["P"] + sp.cost_e * b["E"]) / sp.life_y
                cfg.append(f"园区{p} {b['P']:.0f}kW/{b['E']:.0f}kWh")
                log(f"  独立·园区{p} 最优 {b['P']:.0f}kW/{b['E']:.0f}kWh")
            rows.append({"情景": "独立运营 · 各配最优", "储能配置": "；".join(cfg),
                         "日运行成本": round(ind_best, 2), "年运行成本": round(ind_best * sp.days, 1),
                         "年折旧投资": round(ind_inv, 1),
                         "年化总成本": round(ind_best * sp.days + ind_inv, 1)})
            # 联合
            L, W, V = self.data.joint_arrays(m)
            j0 = op_cost(L, W, V, 0, 0, sp)
            rows.append({"情景": "联合运营 · 无储能", "储能配置": "无",
                         "日运行成本": round(j0, 2), "年运行成本": round(j0 * sp.days, 1),
                         "年折旧投资": 0.0, "年化总成本": round(j0 * sp.days, 1)})
            j50 = op_cost(L, W, V, 50, 100, sp)
            invj50 = (sp.cost_p * 50 + sp.cost_e * 100) / sp.life_y
            rows.append({"情景": "联合运营 · 统一 50/100", "储能配置": "全网 50 kW/100 kWh",
                         "日运行成本": round(j50, 2), "年运行成本": round(j50 * sp.days, 1),
                         "年折旧投资": round(invj50, 1), "年化总成本": round(j50 * sp.days + invj50, 1)})
            rj = optimize_capacity(L, W, V, sp, park="JOINT", P_list=Pl, E_list=El, n_refine=1,
                                   verify=False,
                                   progress=lambda d, t, pk: prog(
                                       int((len(PARKS) + d / max(1, t)) * 1000),
                                       units * 1000, "联合系统"))
            bj = rj["best"]
            rows.append({"情景": "联合运营 · 统一最优", "储能配置": f"全网 {bj['P']:.0f} kW/{bj['E']:.0f} kWh",
                         "日运行成本": round(bj["日运行成本"], 2),
                         "年运行成本": round(bj["年运行成本"], 1),
                         "年折旧投资": round(bj["年折旧投资"], 1),
                         "年化总成本": round(bj["年化总成本"], 1)})
            log(f"联合较独立（无储能）节省 {rows[3]['年化总成本'] - rows[0]['年化总成本']:.1f} 元/年；"
                f"联合最优配置 {bj['P']:.0f}kW/{bj['E']:.0f}kWh")
            return {"view": "joint", "tables": {"联合运营": rows}, "status": "联合运营分析完成"}

        self._run_async(job, "联合运营分析")

    def on_sensitivity(self) -> None:
        if not self._need_data():
            return
        sp = self._spec(); m = self._cur_month()
        parks = PARKS
        scenarios = [
            ("储能单价缩放", "scale_cost", [("0.4×", .4), ("0.6×", .6), ("0.8×", .8),
                                            ("1.0×", 1.0), ("1.2×", 1.2), ("1.5×", 1.5), ("2.0×", 2.0)]),
            ("购电电价", "scale_price_buy", [("0.8", .8), ("0.9", .9), ("1.0", 1.0),
                                             ("1.1", 1.1), ("1.2", 1.2), ("1.4", 1.4)]),
            ("充放电效率", "eta", [("0.85", .85), ("0.88", .88), ("0.90", .90),
                                   ("0.92", .92), ("0.95", .95), ("0.98", .98)]),
            ("运行寿命", "life_y", [("5", 5), ("6", 6), ("8", 8), ("10", 10),
                                    ("12", 12), ("15", 15), ("20", 20)]),
        ]
        # 精度档位：粗/中为纯粗网格（秒级），细加一轮细化（更准但更慢）
        presets = {"coarse": (50.0, 100.0, 0), "mid": (20.0, 50.0, 0), "fine": (10.0, 25.0, 1)}
        step_p, step_e, nref = presets.get(self.cb_sens_prec.currentData(), presets["mid"])
        pmax, emax = self.sp_pmax.value(), self.sp_emax.value()
        Pl, El = build_grid(step_p, step_e, pmax, emax)
        n_levels = sum(len(v) for _, _, v in scenarios)
        n_cfg = n_levels * len(parks)

        def job(log, prog):
            nonlocal Pl, El, nref
            if not HAS_SCIPY and estimate_solves(Pl, El, nref) * n_cfg > 3000:
                # PuLP 每个网格点都要启动一次 CBC 子进程，不自动放粗则基本跑不完
                Pl, El = build_grid(max(step_p, pmax / 4.0), max(step_e, emax / 4.0), pmax, emax)
                nref = 0
                log("⚠ 未检测到 SciPy：网格搜索需反复启动 CBC 子进程（每次约 0.3–0.5 s），"
                    "速度慢约两个数量级，已自动放粗网格。建议 pip install scipy 以获得秒级响应。")
            ns = estimate_solves(Pl, El, nref)
            log(f"⑦ 敏感性分析：{n_levels} 档 × {len(parks)} 园区 = {n_cfg} 个配置，"
                f"每配置 {ns} 次 LP，{solve_budget_note(ns, n_cfg)}")
            rows, t0, k = [], time.time(), 0
            for pname, key, vals in scenarios:
                for lab, v in vals:
                    for p in parks:
                        def sub(d, t, pk, _k=k, _n=pname, _l=lab, _p=p):
                            prog(int((_k + d / max(1, t)) * 1000), n_cfg * 1000,
                                 f"{_n} {_l} · 园区{_p}")
                        r = optimize_capacity(*self._arrays(p, m), sp=sp, park=p,
                                              P_list=Pl, E_list=El, n_refine=nref,
                                              verify=False, progress=sub, **{key: v})
                        b = r["best"]
                        rows.append({"参数": pname, "取值": lab, "园区": p,
                                     "最优P": round(b["P"], 2), "最优E": round(b["E"], 2),
                                     "年化总成本": round(b["年化总成本"], 1)})
                        k += 1
                    log(f"  {pname} {lab} 完成（{k}/{n_cfg}，已用 {time.time() - t0:.0f}s）")
            log(f"⑦ 敏感性分析完成，共 {n_cfg} 个配置，用时 {time.time() - t0:.0f} 秒。")
            return {"view": "sens", "tables": {"敏感性": rows}, "status": "敏感性分析完成"}

        self._run_async(job, "敏感性分析")

    def on_validate(self) -> None:
        if not self._need_data():
            return
        sp = self._spec(); m = self._cur_month(); park = self._cur_park()
        P, E = self.sp_P.value(), self.sp_E.value()
        L, W, V = self._arrays(park, m)
        r = solve_operation(L, W, V, P, E, sp, park, milp=self.chk_milp.isChecked())
        rows = validate_dispatch(r, L, W, V, sp)
        # 双求解器交叉验证
        if HAS_SCIPY and HAS_PULP and P > 0 and E > 0:
            rs = fast_solve(L, W, V, P, E, sp)
            if rs is not None and r.status.startswith("Optimal"):
                d = abs(rs.metrics(L, W, V, sp)["总供电成本"] - r.metrics(L, W, V, sp)["总供电成本"])
                rows.append({"检查项": "双求解器交叉验证 (元)", "实测": fnum(d, 6), "容差": "<1e-4",
                             "结论": "通过" if d < 1e-4 else "提示"})
        self.view = "validate"
        self.payload = {"view": "validate", "tables": {"结果自检": rows}, "status": "结果自检完成"}
        self._fill(self.tables["结果自检"], rows)
        self.tabs.setCurrentWidget(self.tables["结果自检"])
        self._render()
        npass = sum(1 for x in rows if x["结论"] == "通过")
        self._log(f"⑧ 结果自检：{npass}/{len(rows)} 项通过。")

    def on_selftest(self) -> None:
        """⑨ 一键回归自检：用题目默认参数跑完整链路，验证模型与论文一致性。"""
        if not self._need_data():
            return

        def job(log, prog):
            log("一键回归自检开始（固定使用内置模板 + 题目默认参数，与当前界面数据无关）…")
            rows = SelfTestRunner(embedded_data()).run(log)
            for r in rows:
                log(f"  [{'√' if r['结论'] == '通过' else '×'}] {r['检查项']}：{r['说明']}")
            npass = sum(1 for r in rows if r["结论"] == "通过")
            log(f"回归自检结束：{npass}/{len(rows)} 项通过。")
            return {"view": "validate", "tables": {"结果自检": rows},
                    "focus": "结果自检", "status": f"回归自检 {npass}/{len(rows)} 通过"}

        self._run_async(job, "回归自检")

    def on_export(self) -> None:
        if not self.payload.get("tables"):
            QMessageBox.information(self, "提示", "请先运行至少一步计算，再导出。")
            return
        default = _osp.join(os.getcwd(), "microgrid_报表")
        outdir = QFileDialog.getExistingDirectory(self, "选择报表输出目录", default)
        if not outdir:
            return
        try:
            rw = ReportWriter(outdir)
            files = []
            for name, rows in self.payload.get("tables", {}).items():
                p = rw.write_csv(name, rows)
                if p:
                    files.append(p)
            if HAS_MPL and self.fig is not None:
                p = rw.write_png(self.fig)
                if p:
                    files.append(p)
            lines = [f"{APP_NAME} {APP_VERSION}  导出时间 {_dt.datetime.now():%Y-%m-%d %H:%M:%S}", ""]
            sp = self._spec()
            lines.append("当前参数：" + f"园区={self._cur_park()}，数据源={self.cb_src.currentText()}，"
                         f"购电={sp.price_buy}，风电={sp.price_wind}，光伏={sp.price_pv}，"
                         f"售电={sp.price_sell}，分时={sp.tou_enable}，"
                         f"储能 P={self.sp_P.value()}kW E={self.sp_E.value()}kWh，ηc={sp.eta_c}，"
                         f"ηd={sp.eta_d}，SOC=[{sp.soc_min},{sp.soc_max}]，寿命={sp.life_y}年")
            lines.append("")
            for name, rows in self.payload.get("tables", {}).items():
                lines.append(f"—— {name} ——")
                if rows:
                    keys = list(rows[0].keys())
                    lines.append(" | ".join(keys))
                    for r in rows:
                        lines.append(" | ".join(str(r.get(k, "")) for k in keys))
                lines.append("")
            files.append(rw.write_text("报告", lines))
            self._log(f"已导出 {len(files)} 个文件到：{outdir}")
            for f in files:
                self._log("    " + _osp.basename(f))
            QMessageBox.information(self, "导出完成",
                                    f"已生成 {len(files)} 个文件：\n"
                                    + "\n".join(_osp.basename(f) for f in files))
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"{type(e).__name__}: {e}")

    def on_clear(self) -> None:
        self.tab_log.clear()
        for t in self.tables.values():
            self._fill(t, [])
        self.payload = {}
        self.view = "welcome"
        self._render()
        self._log("已清空全部结果与日志。")


# ==============================================================================
#  §8  程序入口
# ==============================================================================
def _install_excepthook(logfile: str) -> None:
    def _hook(etype, value, tb):
        msg = "".join(traceback.format_exception(etype, value, tb))
        try:
            with open(logfile, "a", encoding="utf-8") as f:
                f.write(f"\n[{_dt.datetime.now():%Y-%m-%d %H:%M:%S}] 未捕获异常\n{msg}\n")
        except Exception:
            pass
        sys.__excepthook__(etype, value, tb)

    sys.excepthook = _hook


def run_cli_selftest() -> int:
    """
    命令行自检模式（--selftest）：不启动界面，跑内置回归自检，
    结论同时打印并写入 selftest_report.txt。
    用于验证"打包后的 exe"或"新机器"上环境是否齐备、模型数值是否正确。
    """
    out: List[str] = []

    def log(s: str) -> None:
        out.append(str(s))

    log(f"{APP_NAME} {APP_VERSION} —— 命令行自检")
    log(f"Python {sys.version.split()[0]}｜frozen={getattr(sys, 'frozen', False)}")
    log(f"可执行文件：{sys.executable}")
    try:
        log(f"numpy {np.__version__}")
    except Exception:
        pass
    for ln in engine_report():
        log(ln)
    log("程序图标：" + (icon_path() or "未找到 " + ICON_NAME))
    rows: List[Dict[str, Any]] = []
    try:
        rows = SelfTestRunner(embedded_data()).run(log)
    except Exception:
        log("自检异常：\n" + traceback.format_exc())

    # ---- GUI 栈自检（离屏）：确认 PyQt5 + Matplotlib 画布在打包环境中真的能构建 ----
    try:
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        _app = QApplication.instance() or QApplication(sys.argv[:1])
        win = MainWindow()
        win.resize(1500, 940)
        win.showNormal()
        _app.processEvents()
        win.on_data()
        _app.processEvents()
        win.on_baseline()
        _app.processEvents()
        nax = len(win.fig.axes) if (HAS_MPL and win.fig is not None) else 0
        ok_gui = (win.data is not None) and nax >= 1
        title = win.windowTitle()
        log(f"GUI 栈自检：{'通过' if ok_gui else '未通过'}（{title}，图形轴 {nax} 个）")
        rows.append({"检查项": "GUI 栈（PyQt5 + 绘图画布）",
                     "结论": "通过" if ok_gui else "未通过",
                     "说明": f"{title}，图形轴 {nax} 个"})
        win.close()
    except Exception as e:
        log("GUI 栈自检失败：" + f"{type(e).__name__}: {e}")
        rows.append({"检查项": "GUI 栈（PyQt5 + 绘图画布）", "结论": "未通过",
                     "说明": f"{type(e).__name__}: {str(e)[:120]}"})

    npass = sum(1 for r in rows if r.get("结论") == "通过")
    log("")
    for r in rows:
        log(f"[{'通过' if r['结论'] == '通过' else '未通过'}] {r['检查项']} — {r['说明']}")
    log("")
    log(f"合计 {npass}/{len(rows)} 项通过")
    text = "\n".join(out)
    try:
        print(text)
    except Exception:
        pass
    try:
        base = _osp.dirname(_osp.abspath(sys.executable if getattr(sys, "frozen", False)
                                         else __file__))
        with open(_osp.join(base, "selftest_report.txt"), "w", encoding="utf-8") as f:
            f.write(text + "\n")
    except Exception:
        pass
    return 0 if (rows and npass == len(rows)) else 1


def main() -> int:
    if "--selftest" in sys.argv:
        return run_cli_selftest()
    _install_excepthook(_osp.join(os.getcwd(), "microgrid_System_error.log"))
    try:
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    except Exception:
        pass
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 9))
    ic = icon_path()
    if ic:                                   # 窗口/任务栏图标（exe 文件图标由打包时的 --icon 指定）
        app.setWindowIcon(QIcon(ic))
    win = MainWindow()
    if ic:
        win.setWindowIcon(QIcon(ic))
    win.show()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
