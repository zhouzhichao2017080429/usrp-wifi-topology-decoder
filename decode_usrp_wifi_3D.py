#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
作者：zhouzhichao
日期：2026-06-13

脚本功能：
    默认从脚本同目录或上一级目录 adc signal 文件夹中的所有 .bin 信号文件解析
    USRP B210 采集的 sc16 IQ 数据和 802.11a/g legacy OFDM WiFi 包，
    提取 MAC 地址、SSID、源节点、目的节点、帧类型、FCS 校验等信息，
    最终生成 output.xlsx，并同步生成 index.html 拓扑可视化页面。

流程思路：
    1. 扫描 adc signal 文件夹下所有 .bin 文件，也可以用 --capture 指定单个文件。
    2. 逐个读取 USRP 保存的 sc16 复数 IQ 采样，并做去直流和幅度归一化。
    3. 在每个完整信号中按时间顺序寻找 WiFi 包的起始位置；默认做全信号扫描，
       只有显式指定 --use-hints 或 --hints 时才利用旧 output.xlsx 加速重解码。
    4. 对每个候选包做粗频偏、细频偏估计与校正，再用 L-LTF 做信道估计。
    5. 恢复 L-SIG 字段，得到候选的 MCS 和 PSDU 长度。
    6. 对数据 OFDM 符号做均衡、软解调、去交织、去打孔和 Viterbi 译码。
    7. 解扰得到 MPDU，校验 FCS，并解析 MAC 帧中的节点与通信关系。
    8. 将所有文件的解析结果汇总写入 output.xlsx，并生成节点拓扑图。
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import time
import unicodedata
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
from xml.sax.saxutils import escape

import numpy as np

ROUTER_ICON_NAME = "router_2p5d_clean_node.png"
LAPTOP_ICON_NAME = "laptop_2p5d_intel_node.png"

# Intel 网卡 OUI 前缀库，内置在脚本中，生成拓扑时离线判断 MAC 厂商，不需要联网查询。
INTEL_OUI_PREFIX_TEXT = """
00:02:b3 00:03:47 00:04:23 00:07:e9 00:0c:f1 00:0e:0c 00:0e:35 00:11:11 00:11:75 00:12:f0
00:13:02 00:13:20 00:13:ce 00:13:e8 00:15:00 00:15:17 00:16:6f 00:16:76 00:16:ea 00:16:eb
00:18:de 00:19:d1 00:19:d2 00:1b:21 00:1b:77 00:1c:bf 00:1c:c0 00:1d:e0 00:1d:e1 00:1e:64
00:1e:65 00:1e:67 00:1f:3b 00:1f:3c 00:20:7b 00:21:5c 00:21:5d 00:21:6a 00:21:6b 00:22:fa
00:22:fb 00:23:14 00:23:15 00:24:d6 00:24:d7 00:26:c6 00:26:c7 00:27:0e 00:27:10 00:28:f8
00:42:38 00:72:ee 00:90:27 00:91:9e 00:93:37 00:a0:c9 00:a5:54 00:aa:00 00:aa:01 00:aa:02
00:bb:60 00:c2:c6 00:d0:b7 00:d4:9e 00:d7:6d 00:db:df 00:e1:8c 04:1c:6c 04:33:c2 04:56:e5
04:6c:59 04:a6:c8 04:cf:4b 04:d3:b0 04:e8:b9 04:ea:56 04:ec:d8 04:ed:33 04:f0:ee 08:11:96
08:5b:d6 08:6a:c5 08:71:90 08:8e:90 08:9d:f4 08:b4:d2 08:d2:3e 08:d4:0c 08:eb:21 0c:54:15
0c:7a:15 0c:8b:fd 0c:91:92 0c:9a:3c 0c:ae:39 0c:c9:8a 0c:d2:92 0c:dd:24 10:02:b5 10:0b:a9
10:2e:00 10:3d:1c 10:4a:7d 10:51:07 10:5f:ad 10:91:d1 10:9a:ba 10:a5:1d 10:a8:79 10:f0:05
10:f6:0a 14:18:c3 14:3e:c2 14:4f:8a 14:75:5b 14:85:7f 14:ab:c5 14:f6:d8 18:1d:ea 18:26:49
18:3d:a2 18:56:80 18:5e:0f 18:93:41 18:cc:18 18:ff:0f 1c:1b:b5 1c:4d:70 1c:99:57 1c:c1:0c
20:16:b9 20:1e:88 20:3a:43 20:79:18 20:bd:1d 20:c1:9b 24:41:8c 24:77:03 24:eb:16 24:ee:9a
28:0c:50 28:11:a8 28:16:ad 28:6b:35 28:7f:cf 28:92:00 28:95:29 28:a0:6b 28:a4:4a 28:b2:bd
28:c5:d2 28:c6:3f 28:d0:ea 28:df:eb 2c:0d:a7 2c:33:58 2c:6d:c1 2c:6e:85 2c:7b:a0 2c:8d:b1
2c:db:07 2c:ea:fc 30:05:05 30:24:32 30:3a:64 30:3e:a7 30:89:4a 30:e3:7a 30:e3:a4 30:f6:ef
34:02:86 34:13:e8 34:2e:b7 34:41:5d 34:7d:f6 34:c9:3d 34:cf:f6 34:de:1a 34:e1:2d 34:e6:ad
34:f3:9a 34:f6:4b 34:fd:70 38:00:25 38:18:68 38:68:93 38:7a:0e 38:87:d5 38:ba:f8 38:de:ad
38:fc:98 3c:21:9c 3c:58:c2 3c:6a:a7 3c:9c:0f 3c:a9:f4 3c:e9:f7 3c:f0:11 3c:f8:62 3c:fd:fe
40:1c:83 40:25:c2 40:74:e0 40:a3:cc 40:a6:b7 40:c7:3c 40:d1:33 40:ec:99 40:ec:bd 44:03:2c
44:38:e8 44:49:88 44:85:00 44:a3:bb 44:af:28 44:e5:17 48:40:d5 48:45:20 48:51:b7 48:51:c5
48:68:4a 48:89:e7 48:a4:72 48:ad:9a 48:e1:50 48:f1:7f 4c:03:4f 4c:0f:3e 4c:1d:96 4c:34:88
4c:44:5b 4c:49:6c 4c:5f:70 4c:77:cb 4c:79:6e 4c:79:ba 4c:80:93 4c:a9:54 4c:b0:4a 4c:eb:42
50:28:4a 50:2d:a2 50:2f:9b 50:76:af 50:7c:6f 50:84:92 50:e0:85 50:eb:71 54:14:f3 54:36:31
54:6c:eb 54:8d:5a 54:e4:ed 58:1c:f8 58:6c:25 58:6d:67 58:91:cf 58:94:6b 58:96:1d 58:a0:23
58:a8:39 58:ce:2a 58:fb:84 5c:51:4f 5c:5f:67 5c:67:83 5c:80:b6 5c:87:9c 5c:b2:6d 5c:b4:7e
5c:c5:d4 5c:cd:5b 5c:d2:e4 5c:e0:c5 5c:e4:2a 60:36:dd 60:45:2e 60:57:18 60:67:20 60:6c:66
60:a5:e2 60:dd:8e 60:e3:2b 60:f2:62 60:f6:77 64:32:a8 64:49:7d 64:4a:7d 64:4c:36 64:57:ba
64:5d:86 64:6e:e0 64:79:f0 64:80:99 64:bc:58 64:d4:da 64:d6:9a 64:de:6d 68:05:ca 68:07:15
68:17:29 68:34:21 68:3e:26 68:54:5a 68:5d:43 68:7a:64 68:c6:ac 68:ec:c5 68:f9:0f 6c:29:95
6c:2f:80 6c:4c:e2 6c:6a:77 6c:88:14 6c:94:66 6c:a1:00 6c:f6:da 6c:fe:54 70:08:10 70:15:fb
70:1a:b8 70:1c:e7 70:32:17 70:9c:d1 70:a0:4b 70:a6:cc 70:a8:d3 70:c2:88 70:cd:0d 70:cf:49
70:d8:23 70:d8:c2 74:04:f1 74:13:ea 74:3a:f4 74:70:fd 74:d8:3e 74:e5:0b 74:e5:f9 78:0c:b8
78:2b:46 78:92:9c 78:af:08 78:ff:57 7c:21:4a 7c:2a:31 7c:50:79 7c:5c:f8 7c:67:a2 7c:70:db
7c:76:35 7c:7a:91 7c:b0:c2 7c:b2:7d 7c:b5:66 7c:cc:b8 80:00:0b 80:13:16 80:19:34 80:32:53
80:38:fb 80:45:dd 80:84:89 80:86:f2 80:9b:20 80:b6:55 80:c0:1e 80:e4:ba 84:08:3a 84:14:4d
84:1b:77 84:3a:4b 84:5c:f3 84:68:3e 84:7b:57 84:92:65 84:a6:c8 84:c5:a6 84:d1:c1 84:ef:18
84:fd:d1 88:53:2e 88:78:73 88:b1:11 88:d8:2e 88:f4:da 8c:17:59 8c:1d:96 8c:1f:64 8c:55:4a
8c:70:5a 8c:8d:28 8c:a9:82 8c:b8:7e 8c:c6:81 8c:e9:ee 8c:f8:c5 90:09:df 90:10:57 90:2e:1c
90:47:c2 90:49:fa 90:61:ae 90:65:84 90:78:41 90:b0:21 90:b1:76 90:cc:df 90:e2:ba 94:27:0e
94:39:0e 94:53:ff 94:65:9c 94:b6:09 94:b8:6d 94:e2:3c 94:e6:f7 94:e7:0b 98:2c:bc 98:3b:8f
98:43:fa 98:4f:ee 98:54:1b 98:59:7a 98:5f:41 98:8d:46 98:af:65 98:bd:80 98:fe:3e 9c:29:76
9c:4e:36 9c:65:eb 9c:67:d6 9c:97:1b 9c:b1:50 9c:da:3e 9c:fc:e8 a0:02:a5 a0:29:42 a0:36:9f
a0:4f:52 a0:51:0b a0:59:50 a0:80:69 a0:85:27 a0:88:69 a0:88:b4 a0:a4:c5 a0:a8:cd a0:af:bd
a0:b3:39 a0:c5:89 a0:d3:65 a0:d3:7a a0:e7:0b a4:02:b9 a4:34:d9 a4:42:3b a4:4e:31 a4:6b:b6
a4:b1:c1 a4:bf:01 a4:c3:f0 a4:c4:94 a4:f9:33 a8:59:5f a8:64:f1 a8:6d:aa a8:72:4d a8:7e:ea
ac:05:c7 ac:12:03 ac:16:de ac:19:8e ac:2b:6e ac:3d:cb ac:45:ef ac:5a:fc ac:67:5d ac:72:89
ac:74:b1 ac:7b:a1 ac:82:47 ac:ed:5c ac:fd:ce b0:35:9f b0:3c:dc b0:47:e9 b0:60:88 b0:7d:64
b0:a4:60 b0:dc:ef b4:0e:de b4:69:21 b4:6b:fc b4:6d:83 b4:83:51 b4:96:91 b4:b6:76 b4:d5:bd
b8:03:05 b8:08:cf b8:81:98 b8:8a:60 b8:9a:2a b8:b8:1e b8:bf:83 b8:f7:75 bc:03:58 bc:09:1b
bc:0f:64 bc:17:b8 bc:38:98 bc:54:2f bc:6e:e2 bc:77:37 bc:a8:a6 bc:cd:99 bc:d2:2c bc:f1:05
bc:f1:71 c0:3c:59 c0:a5:e8 c0:a8:10 c0:b6:f9 c0:b8:83 c4:03:a8 c4:0f:08 c4:23:60 c4:3d:1a
c4:47:4e c4:75:ab c4:85:08 c4:bd:e5 c4:d0:e3 c4:d9:87 c4:ff:99 c8:09:a8 c8:15:4e c8:21:58
c8:34:8e c8:58:b3 c8:58:c0 c8:5e:a9 c8:6e:08 c8:8a:9a c8:95:ce c8:b2:9b c8:cb:9e c8:e2:65
c8:f7:33 cc:15:31 cc:2f:71 cc:3d:82 cc:d9:ac cc:f9:e4 d0:3c:1f d0:57:7b d0:57:7e d0:65:78
d0:7e:35 d0:ab:d5 d0:c6:37 d4:25:8b d4:3b:04 d4:54:8b d4:6d:6d d4:94:a9 d4:ab:61 d4:d2:52
d4:d8:53 d4:e9:8a d4:f3:2d d8:3b:bf d8:bf:42 d8:f2:ca d8:f8:83 d8:fc:93 dc:1b:a1 dc:21:48
dc:21:5c dc:41:a9 dc:45:46 dc:46:28 dc:53:60 dc:71:96 dc:8b:28 dc:90:09 dc:97:ba dc:a9:71
dc:fb:48 e0:2b:e9 e0:2e:0b e0:3a:aa e0:72:56 e0:8f:4c e0:94:67 e0:9d:31 e0:c2:64 e0:c9:32
e0:d0:45 e0:d4:64 e0:d4:e8 e0:d5:5d e0:e2:58 e4:02:9b e4:0d:36 e4:1f:d5 e4:42:a6 e4:4a:e0
e4:5e:37 e4:60:17 e4:70:b8 e4:a4:71 e4:a7:a0 e4:b3:18 e4:c7:67 e4:f8:9c e4:fa:fd e4:fd:45
e8:2a:ea e8:62:be e8:84:a5 e8:b0:c5 e8:b1:fc e8:bf:b8 e8:bf:e1 e8:c8:29 e8:f4:08 ec:4c:8c
ec:63:d7 ec:8e:77 ec:e7:a7 ec:ed:04 ec:f3:3c f0:20:ff f0:42:1c f0:57:a6 f0:77:c3 f0:9e:4a
f0:b2:b9 f0:b6:1e f0:d4:15 f0:d5:bf f4:06:69 f4:26:79 f4:3b:d8 f4:46:37 f4:4e:e3 f4:6d:3f
f4:7b:09 f4:8c:50 f4:96:34 f4:a4:75 f4:b3:01 f4:c8:8a f4:ce:23 f4:d1:08 f8:16:54 f8:34:41
f8:59:71 f8:5e:a0 f8:63:3f f8:94:c2 f8:9e:94 f8:ac:65 f8:b5:4d f8:cf:52 f8:e4:e3 f8:f2:1e
f8:fe:5e fc:44:82 fc:6d:77 fc:77:74 fc:9e:53 fc:b3:aa fc:b3:bc fc:f8:ae
"""
INTEL_OUI_PREFIXES = frozenset(INTEL_OUI_PREFIX_TEXT.split())

# 802.11a/g 20 MHz legacy OFDM 的基础参数：
# 20 Msps 采样率、64 点 FFT、16 点循环前缀、16 点短训练重复周期。
SAMPLE_RATE = 20_000_000.0
FFT_LEN = 64
CP_LEN = 16
SHORT_LEN = 16

# 数据子载波 48 个，导频子载波 4 个；DC 子载波 0 不使用。
DATA_SC = np.array(
    list(range(-26, -21)) + list(range(-20, -7)) + list(range(-6, 0))
    + list(range(1, 7)) + list(range(8, 21)) + list(range(22, 27)),
    dtype=np.int16,
)
PILOT_SC = np.array([-21, -7, 7, 21], dtype=np.int16)

LLTF_53 = np.array([
    1, 1, -1, -1, 1, 1, -1, 1, -1, 1, 1, 1, 1, 1, 1, -1, -1, 1,
    1, -1, 1, -1, 1, 1, 1, 1, 0, 1, -1, -1, 1, 1, -1, 1, -1, 1,
    -1, -1, -1, -1, -1, 1, 1, -1, -1, 1, -1, 1, -1, 1, 1, 1, 1,
], dtype=np.float32)

RATE_TO_MCS = {
    0b1101: 0,
    0b1111: 1,
    0b0101: 2,
    0b0111: 3,
    0b1001: 4,
    0b1011: 5,
    0b0001: 6,
    0b0011: 7,
}

MCS_TABLE = {
    0: {"nbpsc": 1, "ncbps": 48, "ndbps": 24, "rate": "1/2", "mod": "BPSK"},
    1: {"nbpsc": 1, "ncbps": 48, "ndbps": 36, "rate": "3/4", "mod": "BPSK"},
    2: {"nbpsc": 2, "ncbps": 96, "ndbps": 48, "rate": "1/2", "mod": "QPSK"},
    3: {"nbpsc": 2, "ncbps": 96, "ndbps": 72, "rate": "3/4", "mod": "QPSK"},
    4: {"nbpsc": 4, "ncbps": 192, "ndbps": 96, "rate": "1/2", "mod": "16QAM"},
    5: {"nbpsc": 4, "ncbps": 192, "ndbps": 144, "rate": "3/4", "mod": "16QAM"},
    6: {"nbpsc": 6, "ncbps": 288, "ndbps": 192, "rate": "2/3", "mod": "64QAM"},
    7: {"nbpsc": 6, "ncbps": 288, "ndbps": 216, "rate": "3/4", "mod": "64QAM"},
}

HEADERS = [
    "序号", "SSID名称", "BSSID", "帧大类", "帧子类型",
    "源节点", "目的节点", "通信方向", "接收节点RA", "发送节点TA",
    "地址1", "地址2", "地址3", "地址4", "ToDS", "FromDS",
    "FCS校验通过", "起始采样点", "MCS", "PSDU长度", "比特顺序",
    "频偏Hz", "序列号", "片段号", "Duration", "采集数据文件",
]

# 解码参数集合：控制搜索速度、候选范围、提示表、最大包数等行为。
@dataclass
class DecodeOptions:
    fast_mode: bool = True
    try_both_byte_orders: bool = False
    exhaustive_mcs: bool = False
    verbose_frames: bool = False
    progress_step_percent: float = 5.0
    detect_window_samples: int = 40_000
    detect_overlap_samples: int = 4_000
    energy_gate_factor: float = 3.0
    lltf_max_candidates: int = 64
    stf_max_candidates: int = 64
    fallback_on_miss: bool = False
    stop_after_fcs_ok: bool = True
    max_complex_samples: Optional[int] = None
    max_packets: Optional[int] = None
    max_frames: Optional[int] = None
    max_psdu_length: int = 800
    scan_step_samples: int = 320
    rescue_control_frames: bool = True

# 单个 MAC 帧的解析结果：保存地址字段、帧类型、SSID、源节点和目的节点。
@dataclass
class MacInfo:
    valid: bool = False
    type_num: int = -1
    subtype_num: int = -1
    type_name: str = "unknown"
    subtype_name: str = "unknown"
    to_ds: int = 0
    from_ds: int = 0
    duration: Optional[int] = None
    addr1: str = ""
    addr2: str = ""
    addr3: str = ""
    addr4: str = ""
    ra: str = ""
    ta: str = ""
    source_node: str = ""
    destination_node: str = ""
    bssid: str = ""
    ssid: str = ""
    sequence_number: Optional[int] = None
    fragment_number: Optional[int] = None

# 最终写入 Excel 的一行结果：把物理层解码信息和 MAC 层节点信息合在一起。
@dataclass
class ResultEntry:
    ssid: str
    bssid: str
    frame_type: str
    frame_subtype: str
    source_node: str
    destination_node: str
    direction: str
    receiver_addr: str
    transmitter_addr: str
    addr1: str
    addr2: str
    addr3: str
    addr4: str
    to_ds: int
    from_ds: int
    fcs_ok: int
    start_sample: int
    mcs: int
    psdu_length: int
    bit_order: str
    cfo_hz: float
    sequence_number: Optional[int]
    fragment_number: Optional[int]
    duration: Optional[int]

# 多文件批量解码时，保存每一帧结果和它来自哪个 .bin 采集文件。
@dataclass
class CaptureResult:
    capture_file: Path
    result: ResultEntry

# 从旧 output.xlsx 读取到的提示项：用于按已知起始点、MCS 和 PSDU 长度快速重解码。
@dataclass
class HintEntry:
    start_sample: int
    mcs: Optional[int] = None
    psdu_length: Optional[int] = None

def env_flag(name: str, default: bool) -> bool:
    text = os.environ.get(name, "").strip().lower()
    if not text:
        return default
    if text in {"1", "true", "yes", "on"}:
        return True
    if text in {"0", "false", "no", "off"}:
        return False
    return default

def env_number(name: str, default: Optional[float]) -> Optional[float]:
    text = os.environ.get(name, "").strip()
    if not text:
        return default
    if text.lower() in {"inf", "infinity"}:
        return None
    try:
        return float(text)
    except ValueError:
        return default

# 确定待解析的 USRP 采集文件：优先用命令行/环境变量，其次搜索脚本附近的 .bin 文件。
def resolve_capture(script_dir: Path, preferred: str) -> Path:
    override = os.environ.get("WIFI_CAPTURE_FILE", "").strip()
    if override:
        path = Path(override)
        if not path.exists():
            raise FileNotFoundError(f"WIFI_CAPTURE_FILE does not exist: {path}")
        return path

    candidates = [script_dir, script_dir.parent / "auto code"]
    for folder in candidates:
        path = folder / preferred
        if path.exists():
            return path
    for folder in candidates:
        bins = sorted(folder.glob("*.bin"))
        if bins:
            return bins[0]
    raise FileNotFoundError(f"No .bin USRP IQ capture was found in {script_dir}")

# 扫描默认 adc signal 文件夹中的所有正式 .bin 文件；默认会跳过 debug/test/seq 等临时文件。
def resolve_captures(script_dir: Path, capture: str = "", capture_dir: str = "") -> List[Path]:
    if capture:
        path = Path(capture)
        if not path.exists():
            raise FileNotFoundError(f"Capture file does not exist: {path}")
        return [path]
    folder = Path(capture_dir) if capture_dir else script_dir / "adc signal"
    if not folder.exists():
        raise FileNotFoundError(f"Capture directory does not exist: {folder}")
    skipped_prefixes = ("debug_", "test_", "seq")
    captures = [
        path for path in sorted(folder.glob("*.bin"))
        if path.is_file() and not path.name.lower().startswith(skipped_prefixes)
    ]
    if not captures:
        raise FileNotFoundError(f"No .bin USRP IQ captures were found in {folder}")
    return captures

# 读取 USRP sc16 IQ 文件，并转换为归一化后的复数采样序列。
def read_capture(path: Path, fmt: str, max_complex: Optional[int]) -> np.ndarray:
    if fmt.lower() != "sc16":
        raise ValueError("Only sc16 capture format is implemented in this Python version")
    count = -1 if max_complex is None else int(max_complex) * 2
    raw = np.fromfile(path, dtype=np.int16, count=count).astype(np.float32) / 32768.0
    if raw.size < 20_000:
        raise ValueError(f"Too few raw samples in capture: {raw.size}")
    if raw.size % 2:
        raw = raw[:-1]
    rx = raw[0::2] + 1j * raw[1::2]
    rx = rx.astype(np.complex64)
    rx -= np.mean(rx)
    peak = np.max(np.abs(rx))
    if peak > 0:
        rx /= peak
    return rx

# 快速抽样判断文件是否可能包含明显 WiFi 活动；空信道可直接跳过完整解码。
def capture_has_activity(path: Path, factor: float = 8.0, max_int16: int = 1_000_000) -> bool:
    if factor <= 0:
        return True
    raw = np.fromfile(path, dtype=np.int16, count=max_int16)
    if raw.size < 20_000:
        return True
    if raw.size % 2:
        raw = raw[:-1]
    i = raw[0::2].astype(np.float32)
    q = raw[1::2].astype(np.float32)
    power = i * i + q * q
    if power.size == 0:
        return True
    step = max(1, power.size // 50_000)
    sampled = power[::step]
    floor = max(float(np.median(sampled)), 1.0)
    peak = float(np.percentile(sampled, 99.9))
    return peak > floor * factor

def sc_to_fft_index(sc: int) -> int:
    return sc % FFT_LEN

# 构造 L-LTF 在频域上的已知训练序列。
def lltf_freq() -> np.ndarray:
    freq = np.zeros(FFT_LEN, dtype=np.complex64)
    subcarriers = list(range(-26, 27))
    for sc, val in zip(subcarriers, LLTF_53):
        if sc != 0:
            freq[sc_to_fft_index(sc)] = val
    return freq

LLTF_FREQ = lltf_freq()
LLTF_TIME = np.fft.ifft(LLTF_FREQ).astype(np.complex64)
LLTF_FIELD = np.concatenate([LLTF_TIME[-32:], LLTF_TIME, LLTF_TIME]).astype(np.complex64)

# 用窗口能量判断是否值得做昂贵的 L-LTF/STF 相关；空白信道会被快速跳过。
def window_has_signal(window: np.ndarray, noise_floor: float, factor: float) -> bool:
    if factor <= 0 or window.size < 320:
        return True
    step = max(80, min(800, window.size // 200))
    sampled_power = np.abs(window[::step]) ** 2
    if sampled_power.size == 0:
        return True
    peak = float(np.max(sampled_power))
    local_floor = max(noise_floor, float(np.median(sampled_power)) * 1.5, 1e-8)
    return peak > local_floor * factor

# 用短训练字段的周期相关性，在一个信号片段中寻找多个可能的 WiFi 包起始点。
def detect_packet_offsets(seg: np.ndarray, threshold: float = 0.72, max_candidates: int = 32) -> List[int]:
    if seg.size < 320:
        return []
    a = seg[:-SHORT_LEN]
    b = seg[SHORT_LEN:]
    prod = np.conj(a) * b
    p = np.convolve(prod, np.ones(SHORT_LEN, dtype=np.complex64), mode="valid")
    r = np.convolve(np.abs(b) ** 2, np.ones(SHORT_LEN, dtype=np.float32), mode="valid")
    metric = (np.abs(p) ** 2) / (r ** 2 + 1e-12)
    power = r / SHORT_LEN
    floor = max(float(np.median(power)) * 2.0, 2e-5)
    idx = np.flatnonzero((metric > threshold) & (power > floor))
    if idx.size == 0:
        return []

    offsets: List[int] = []
    group_start = int(idx[0])
    previous = int(idx[0])
    for value in idx[1:]:
        value = int(value)
        if value - previous > 24:
            offsets.append(max(0, group_start - 16))
            if len(offsets) >= max_candidates:
                return offsets
            group_start = value
        previous = value
    offsets.append(max(0, group_start - 16))
    return offsets[:max_candidates]

# 用 L-LTF 相关峰对粗略包起点做细化，提高后续 FFT 和信道估计稳定性。
def refine_packet_start(rx: np.ndarray, coarse_start: int) -> int:
    search0 = max(0, coarse_start - 48)
    search1 = min(rx.size - 320, coarse_start + 80)
    if search1 <= search0:
        return coarse_start
    ref = LLTF_FIELD
    seg = rx[search0 + 160:search1 + 160 + ref.size - 1]
    if seg.size < ref.size:
        return coarse_start
    corr = np.abs(np.convolve(seg, np.conj(ref[::-1]), mode="valid"))
    if corr.size == 0:
        return coarse_start
    return search0 + int(np.argmax(corr))

# 直接用 L-LTF 模板相关搜索候选起点，作为包检测的主要候选来源之一。
def lltf_candidate_offsets(seg: np.ndarray, max_candidates: int = 24, earliest_only: bool = False) -> List[int]:
    if seg.size < 480:
        return []
    ref = LLTF_FIELD
    corr = np.abs(np.convolve(seg, np.conj(ref[::-1]), mode="valid"))
    energy = np.convolve(np.abs(seg) ** 2, np.ones(ref.size, dtype=np.float32), mode="valid")
    metric = corr / (np.sqrt(energy) * (float(np.linalg.norm(ref)) + 1e-12) + 1e-12)
    if metric.size <= 160:
        return []
    noise = float(np.median(metric))
    threshold = max(0.30, noise * 8.0)
    peaks = np.flatnonzero(metric > threshold)
    if peaks.size == 0:
        return []
    order = peaks if earliest_only else peaks[np.argsort(metric[peaks])[::-1]]
    offsets: List[int] = []
    for ltf_start in order:
        start = int(ltf_start) - 160
        if start < 0:
            continue
        if any(abs(start - old) < 240 for old in offsets):
            continue
        offsets.append(start)
        if len(offsets) >= max_candidates:
            break
    return sorted(offsets)

# 根据短训练字段估计粗频偏。
def estimate_coarse_cfo(stf: np.ndarray) -> float:
    if stf.size < 160:
        return 0.0
    p = np.vdot(stf[:64], stf[64:128])
    return float(np.angle(p) * SAMPLE_RATE / (2 * np.pi * 64))

# 根据两个长训练符号之间的相位差估计细频偏。
def estimate_fine_cfo(lltf_field: np.ndarray) -> float:
    if lltf_field.size < 160:
        return 0.0
    sym1 = lltf_field[24:88]
    sym2 = lltf_field[88:152]
    p = np.vdot(sym1, sym2)
    return float(np.angle(p) * SAMPLE_RATE / (2 * np.pi * 64))

# 对复数采样做频偏补偿。
def apply_cfo(x: np.ndarray, cfo_hz: float, sample_offset: int = 0) -> np.ndarray:
    n = np.arange(x.size, dtype=np.float32) + sample_offset
    return x * np.exp(-1j * 2 * np.pi * cfo_hz * n / SAMPLE_RATE).astype(np.complex64)

# 使用 L-LTF 估计所有有效子载波、数据子载波和导频子载波的信道响应。
def estimate_channel(pkt: np.ndarray) -> Tuple[np.ndarray, np.ndarray]:
    lltf = pkt[160:320]
    sym1 = np.fft.fft(lltf[32:96])
    sym2 = np.fft.fft(lltf[96:160])
    avg = (sym1 + sym2) / 2
    h = np.zeros(FFT_LEN, dtype=np.complex64)
    nonzero = np.abs(LLTF_FREQ) > 0
    h[nonzero] = avg[nonzero] / LLTF_FREQ[nonzero]
    data_h = np.array([h[sc_to_fft_index(sc)] for sc in DATA_SC], dtype=np.complex64)
    pilot_h = np.array([h[sc_to_fft_index(sc)] for sc in PILOT_SC], dtype=np.complex64)
    return data_h, pilot_h

# 从候选包起点估计频偏和信道，返回总频偏、校正后的前导、数据/导频信道。
def estimate_frontend(rx: np.ndarray, pkt_start: int) -> Tuple[float, np.ndarray, np.ndarray, np.ndarray]:
    pre = rx[pkt_start:pkt_start + 400]
    coarse = estimate_coarse_cfo(pre[:160])
    pre_c = apply_cfo(pre, coarse, 0)
    cfo = coarse + estimate_fine_cfo(pre_c[160:320])
    pre_cf = apply_cfo(pre, cfo, 0)
    data_h, pilot_h = estimate_channel(pre_cf)
    return cfo, pre_cf, data_h, pilot_h

# 按 IEEE 802.11 规则对一个 OFDM 符号内的软比特做去交织。
def deinterleave(values: np.ndarray, ncbps: int, nbpsc: int) -> np.ndarray:
    s = max(nbpsc // 2, 1)
    out = np.empty_like(values)
    for k in range(ncbps):
        i = (ncbps // 16) * (k % 16) + k // 16
        j = s * (i // s) + (i + ncbps - (16 * i) // ncbps) % s
        out[k] = values[j]
    return out

# 计算卷积编码器在给定状态和输入比特下输出的两个编码比特。
def conv_outputs(state: int, bit: int, polys: Tuple[int, int] = (0o133, 0o171)) -> Tuple[int, int]:
    # IEEE 802.11 卷积编码器把最新输入比特放在 7 bit 移位寄存器的高位。
    # 这个方向和 MATLAB WLAN Toolbox 对参考采集文件恢复出的 L-SIG 比特一致。
    reg = (state | ((int(bit) & 1) << 6)) & 0x7F
    g0, g1 = polys
    o0 = bin(reg & g0).count("1") & 1
    o1 = bin(reg & g1).count("1") & 1
    return o0, o1

def conv_next_state(state: int, bit: int) -> int:
    return ((state >> 1) | ((int(bit) & 1) << 5)) & 0x3F

# 对接收到的软比特补回被打孔的位置，补入 0 作为中性软信息。
def depuncture(soft: np.ndarray, rate: str) -> np.ndarray:
    if rate == "1/2":
        return soft
    if rate == "2/3":
        pattern = np.array([1, 1, 1, 0], dtype=np.uint8)
    elif rate == "3/4":
        pattern = np.array([1, 1, 1, 0, 0, 1], dtype=np.uint8)
    else:
        return soft
    n_full = int(math.ceil(soft.size * pattern.size / float(np.sum(pattern))))
    out: List[float] = []
    src = 0
    p = 0
    while src < soft.size:
        if pattern[p % pattern.size]:
            out.append(float(soft[src]))
            src += 1
        else:
            out.append(0.0)
        p += 1
    if len(out) % 2:
        out.append(0.0)
    return np.array(out[:n_full + (n_full % 2)], dtype=np.float32)

# 返回 legacy OFDM 第 symbol_index 个数据符号对应的导频极性。
def legacy_pilot_polarity(symbol_index: int) -> int:
    # 802.11a/g 标准规定的导频极性序列，从第一个数据符号开始编号。
    seq = [
        1, 1, 1, 1, -1, -1, -1, 1, -1, -1, -1, -1, 1, 1, -1, 1,
        -1, -1, 1, 1, -1, 1, 1, -1, 1, 1, 1, 1, 1, 1, -1, 1,
        1, 1, -1, 1, 1, -1, -1, 1, 1, 1, -1, 1, -1, -1, -1, 1,
        -1, 1, -1, -1, 1, -1, -1, 1, 1, 1, 1, 1, -1, -1, 1, 1,
        -1, -1, 1, -1, 1, -1, 1, 1, -1, -1, -1, 1, 1, -1, -1, -1,
        -1, 1, -1, -1, 1, -1, 1, 1, 1, 1, -1, 1, -1, 1, -1, 1,
        -1, -1, -1, -1, -1, 1, -1, 1, 1, -1, 1, -1, 1, 1, 1, -1,
        -1, 1, -1, -1, -1, 1, 1, 1, -1, -1, -1, -1, -1, -1, -1,
    ]
    return seq[symbol_index % len(seq)]

# 对一个 OFDM 符号做信道均衡，并利用导频估计和补偿公共相位误差。
def equalize_symbol(
    sym: np.ndarray,
    data_h: np.ndarray,
    pilot_h: np.ndarray,
    symbol_index: int,
    cp_offset: int = 16,
    cpe_mode: int = 0,
) -> np.ndarray:
    freq = np.fft.fft(sym[cp_offset:cp_offset + FFT_LEN])
    data = np.array([freq[sc_to_fft_index(sc)] for sc in DATA_SC], dtype=np.complex64)
    pilots = np.array([freq[sc_to_fft_index(sc)] for sc in PILOT_SC], dtype=np.complex64)
    eq = data / (data_h + 1e-12)
    peq = pilots / (pilot_h + 1e-12)
    pilot_ref = legacy_pilot_polarity(symbol_index) * np.array([1, 1, 1, -1], dtype=np.complex64)
    # 估计公共相位误差；这里的负号和 MATLAB WLAN Toolbox 在该采集上的极性一致。
    raw_cpe = np.angle(np.sum(peq * np.conj(-pilot_ref)))
    if cpe_mode == 1:
        cpe = -raw_cpe
    elif cpe_mode == 2:
        cpe = 0.0
    else:
        cpe = raw_cpe
    return eq * np.exp(-1j * cpe)

# 根据 MCS 对均衡后的星座点做软解调，输出给 Viterbi 的软比特。
def demap_soft(eq: np.ndarray, mcs: int) -> np.ndarray:
    if mcs in (0, 1):
        return np.real(eq).astype(np.float32)
    if mcs in (2, 3):
        vals = np.empty(eq.size * 2, dtype=np.float32)
        vals[0::2] = -np.real(eq)
        vals[1::2] = -np.imag(eq)
        return vals
    if mcs in (4, 5):
        scale = np.sqrt(10.0)
        x = eq * scale
        vals = np.empty(eq.size * 4, dtype=np.float32)
        # 16QAM 的第 1/3 个软比特对应 I/Q 符号位。旧版这里符号取反，
        # 会导致 MCS4/5 的 RTS 等短控制帧大量 FCS 失败；按 WLAN Toolbox
        # 对同一采样的恢复结果，应使用正向 I/Q 符号位。
        vals[0::4] = np.real(x)
        vals[1::4] = np.abs(np.real(x)) - 2.0
        vals[2::4] = np.imag(x)
        vals[3::4] = np.abs(np.imag(x)) - 2.0
        return vals
    if mcs in (6, 7):
        scale = np.sqrt(42.0)
        x = eq * scale
        vals = np.empty(eq.size * 6, dtype=np.float32)
        vals[0::6] = -np.real(x)
        vals[1::6] = np.abs(np.real(x)) - 4.0
        vals[2::6] = np.abs(4.0 - np.abs(np.real(x))) - 2.0
        vals[3::6] = -np.imag(x)
        vals[4::6] = np.abs(np.imag(x)) - 4.0
        vals[5::6] = np.abs(4.0 - np.abs(np.imag(x))) - 2.0
        return vals
    return np.real(eq).astype(np.float32)

# 对 1/2 码率卷积码软比特做 Viterbi 译码。
def viterbi_decode_rate_half(soft: np.ndarray, polys: Tuple[int, int] = (0o133, 0o171)) -> np.ndarray:
    pairs = soft.reshape(-1, 2)
    n_steps = pairs.shape[0]
    inf = 1e30
    metrics = np.full(64, inf, dtype=np.float64)
    metrics[0] = 0.0
    prev_state = np.zeros((n_steps, 64), dtype=np.uint8)
    prev_bit = np.zeros((n_steps, 64), dtype=np.uint8)

    for t in range(n_steps):
        new_metrics = np.full(64, inf, dtype=np.float64)
        r0, r1 = float(pairs[t, 0]), float(pairs[t, 1])
        for state in range(64):
            base = metrics[state]
            if base >= inf / 2:
                continue
            for bit in (0, 1):
                o0, o1 = conv_outputs(state, bit, polys)
                s0 = 1.0 if o0 == 0 else -1.0
                s1 = 1.0 if o1 == 0 else -1.0
                branch = (r0 - s0) ** 2 + (r1 - s1) ** 2
                ns = conv_next_state(state, bit)
                metric = base + branch
                if metric < new_metrics[ns]:
                    new_metrics[ns] = metric
                    prev_state[t, ns] = state
                    prev_bit[t, ns] = bit
        metrics = new_metrics

    state = 0 if metrics[0] < inf / 2 else int(np.argmin(metrics))
    bits = np.zeros(n_steps, dtype=np.uint8)
    for t in range(n_steps - 1, -1, -1):
        bit = prev_bit[t, state]
        bits[t] = bit
        state = int(prev_state[t, state])
    return bits

def bits_to_int_lsb(bits: Sequence[int]) -> int:
    return int(sum(int(b) << i for i, b in enumerate(bits)))

def bits_to_int_msb(bits: Sequence[int]) -> int:
    value = 0
    for bit in bits:
        value = (value << 1) | int(bit)
    return value

def ratecode_to_mcs(code: int) -> int:
    return RATE_TO_MCS.get(int(code), -1)

# 检查 L-SIG 字段的奇偶校验、尾比特、速率编码和长度是否合理。
def lsig_quality(bits: np.ndarray) -> Tuple[int, int, bool]:
    parity_ok = int((int(np.sum(bits[:17]) & 1) == int(bits[17])))
    tail_zeros = 6 - int(np.sum(bits[18:24]))
    # MATLAB 版本同时尝试 RATE 和 LENGTH 的 LSB/MSB 解释；这里的质量检查也要
    # 按同样规则判定，否则会把 MATLAB 能继续尝试的 L-SIG 候选提前丢掉。
    mcs_ok = int(
        ratecode_to_mcs(bits_to_int_msb(bits[:4])) >= 0
        or ratecode_to_mcs(bits_to_int_lsb(bits[:4])) >= 0
    )
    length_ok = int(
        0 < bits_to_int_lsb(bits[5:17]) <= 4095
        or 0 < bits_to_int_msb(bits[5:17]) <= 4095
    )
    strict = bool(parity_ok and tail_zeros == 6 and mcs_ok and length_ok)
    return parity_ok + tail_zeros + mcs_ok + length_ok, tail_zeros, strict

# 从包前导后的 SIGNAL 字段恢复 24 bit L-SIG 信息。
def recover_lsig_bits(pkt: np.ndarray, data_h: np.ndarray, pilot_h: np.ndarray) -> Optional[np.ndarray]:
    if pkt.size < 400:
        return None
    sig = pkt[320:400]
    candidates: List[Tuple[int, int, np.ndarray]] = []
    for cpe_symbol_index in (0, -1):
        eq = equalize_symbol(sig, data_h, pilot_h, cpe_symbol_index, cp_offset=16)
        for sign in (1.0, -1.0):
            soft = deinterleave((sign * np.real(eq)).astype(np.float32), 48, 1)
            for polys in ((0o133, 0o171), (0o171, 0o133)):
                bits = viterbi_decode_rate_half(soft, polys=polys)
                if bits.size < 24:
                    continue
                bits = bits[:24]
                score, tail, strict = lsig_quality(bits)
                candidates.append((score, tail, bits, strict))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return candidates[0][2] if candidates[0][3] else None

# 从 L-SIG 中生成候选 PSDU 长度，快速模式只取最可能的一种。
def candidate_lengths(bits: np.ndarray, fast: bool) -> List[int]:
    lsb = bits_to_int_lsb(bits[5:17])
    msb = bits_to_int_msb(bits[5:17])
    if fast:
        return [lsb if 0 < lsb <= 4095 else msb]
    return sorted({v for v in (lsb, msb) if 0 < v <= 4095})

def filter_lengths(lengths: Sequence[int], opt: DecodeOptions) -> List[int]:
    return [int(v) for v in lengths if 0 < int(v) <= int(opt.max_psdu_length)]

# 从 L-SIG 速率字段生成候选 MCS，必要时可穷举所有 legacy MCS。
def candidate_mcs(bits: np.ndarray, exhaustive: bool) -> List[int]:
    code_msb = bits_to_int_msb(bits[:4])
    code_lsb = bits_to_int_lsb(bits[:4])
    values = {ratecode_to_mcs(code_msb), ratecode_to_mcs(code_lsb)}
    if exhaustive:
        values.update(range(8))
    return sorted(v for v in values if v >= 0)

def mcs_symbol_count(mcs: int, psdu_length: int) -> int:
    ndbps = MCS_TABLE[mcs]["ndbps"]
    return int(math.ceil((16 + 8 * psdu_length + 6) / ndbps))

# 恢复一个数据包的编码比特流：均衡所有数据符号，再软解调、去交织、去打孔和 Viterbi 译码。
def recover_data_bits(
    pkt: np.ndarray,
    data_h: np.ndarray,
    pilot_h: np.ndarray,
    mcs: int,
    psdu_length: int,
    cpe_mode: int = 0,
) -> Optional[np.ndarray]:
    if mcs not in MCS_TABLE:
        return None
    table = MCS_TABLE[mcs]
    nsym = mcs_symbol_count(mcs, psdu_length)
    needed = 400 + nsym * 80
    if pkt.size < needed:
        return None
    all_soft: List[np.ndarray] = []
    for sidx in range(nsym):
        sym = pkt[400 + sidx * 80:400 + (sidx + 1) * 80]
        eq = equalize_symbol(sym, data_h, pilot_h, sidx + 1, cp_offset=16, cpe_mode=cpe_mode)
        soft = demap_soft(eq, mcs)
        csi = np.repeat((np.abs(data_h) ** 2).astype(np.float32), int(table["nbpsc"]))
        soft = deinterleave(soft * csi, table["ncbps"], table["nbpsc"])
        all_soft.append(soft)
    soft_bits = np.concatenate(all_soft)
    soft_depunctured = depuncture(soft_bits, str(table["rate"]))
    decoded = viterbi_decode_rate_half(soft_depunctured)
    need_bits = 16 + 8 * psdu_length + 6
    if decoded.size < need_bits:
        return None
    return decoded[:need_bits]

# 生成 802.11 扰码序列，用于后续 MPDU 解扰。
def scramble_sequence(seed: int, nbits: int) -> np.ndarray:
    state = [(seed >> i) & 1 for i in range(7)]
    out = np.zeros(nbits, dtype=np.uint8)
    for i in range(nbits):
        feedback = state[3] ^ state[6]
        out[i] = feedback
        state = [feedback] + state[:6]
    return out

# 将比特数组按指定 bit 顺序打包成字节。
def bits_to_bytes(bits: np.ndarray, lsb_first: bool = True) -> bytes:
    nbytes = bits.size // 8
    bits = bits[:nbytes * 8].reshape(nbytes, 8)
    if lsb_first:
        weights = (1 << np.arange(8, dtype=np.uint16))
    else:
        weights = (1 << np.arange(7, -1, -1, dtype=np.uint16))
    vals = (bits.astype(np.uint16) * weights).sum(axis=1).astype(np.uint8)
    return bytes(vals.tolist())

# 尝试所有扰码种子解扰 PSDU，并优先返回 FCS 正确的 MPDU。
def descramble_payload(decoded_bits: np.ndarray, psdu_length: int, lsb_first: bool) -> Optional[bytes]:
    best_payload = None
    for seed in range(1, 128):
        seq = scramble_sequence(seed, decoded_bits.size)
        plain = decoded_bits ^ seq
        if np.any(plain[:7] != 0):
            continue
        payload_bits = plain[16:16 + 8 * psdu_length]
        if payload_bits.size < 8 * psdu_length:
            continue
        payload = bits_to_bytes(payload_bits, lsb_first=lsb_first)
        if check_fcs(payload):
            return payload
        if best_payload is None:
            best_payload = payload
    return best_payload

# 按 802.11 FCS 使用的 CRC32 多项式计算字节流校验值。
def crc32_bytes(data: bytes) -> int:
    crc = 0xFFFFFFFF
    for b in data:
        crc ^= b
        for _ in range(8):
            if crc & 1:
                crc = (crc >> 1) ^ 0xEDB88320
            else:
                crc >>= 1
            crc &= 0xFFFFFFFF
    return (~crc) & 0xFFFFFFFF

# 检查 MPDU 尾部 4 字节 FCS 是否与正文 CRC32 匹配。
def check_fcs(mpdu: bytes) -> bool:
    if len(mpdu) < 8:
        return False
    got = int.from_bytes(mpdu[-4:], "little")
    return got == crc32_bytes(mpdu[:-4])

def mac_at(data: bytes, start0: int) -> str:
    if start0 + 6 > len(data):
        return ""
    return ":".join(f"{b:02x}" for b in data[start0:start0 + 6])

def type_name(typ: int) -> str:
    return ["management", "control", "data", "reserved"][typ] if 0 <= typ <= 3 else "unknown"

def subtype_name(typ: int, sub: int) -> str:
    names = {
        0: [
            "association_request", "association_response", "reassociation_request",
            "reassociation_response", "probe_request", "probe_response",
            "timing_advertisement", "reserved", "beacon", "atim",
            "disassociation", "authentication", "deauthentication", "action",
            "action_no_ack", "reserved",
        ],
        1: [
            "reserved", "reserved", "trigger", "tack", "beamforming_report_poll",
            "vht_ndp_announcement", "control_frame_extension", "control_wrapper",
            "block_ack_request", "block_ack", "ps_poll", "rts", "cts", "ack",
            "cf_end", "cf_end_cf_ack",
        ],
        2: [
            "data", "data_cf_ack", "data_cf_poll", "data_cf_ack_cf_poll", "null",
            "cf_ack", "cf_poll", "cf_ack_cf_poll", "qos_data", "qos_data_cf_ack",
            "qos_data_cf_poll", "qos_data_cf_ack_cf_poll", "qos_null", "reserved",
            "qos_cf_poll", "qos_cf_ack_cf_poll",
        ],
    }
    return names.get(typ, ["reserved"] * 16)[sub] if 0 <= sub < 16 else "unknown"

SSID_ALLOWED_SYMBOLS = set(" -_.:@#[]()+=~!$%^&,")

def is_common_ssid_char(ch: str) -> bool:
    code = ord(ch)
    return (
        ch in SSID_ALLOWED_SYMBOLS
        or 48 <= code <= 57
        or 65 <= code <= 90
        or 97 <= code <= 122
        or 0x3400 <= code <= 0x4DBF
        or 0x4E00 <= code <= 0x9FFF
        or 0xF900 <= code <= 0xFAFF
        or 0x3040 <= code <= 0x30FF
        or 0xAC00 <= code <= 0xD7AF
    )

def clean_ssid_text(text: str) -> str:
    raw = str(text or "")
    if raw.strip() == "<hidden>":
        return "<hidden>"
    raw = re.sub(r"<hidden>", " ", raw, flags=re.IGNORECASE)
    chars: List[str] = []
    for ch in raw.replace("\ufffd", ""):
        category = unicodedata.category(ch)
        if category.startswith("C"):
            continue
        if ch.isspace():
            chars.append(" ")
        elif is_common_ssid_char(ch):
            chars.append(ch)
    cleaned = re.sub(r"\s+", " ", "".join(chars)).strip()
    if not cleaned or not any(ch.isalnum() for ch in cleaned):
        return ""
    return cleaned

def ssid_candidate_score(raw_text: str, cleaned: str) -> int:
    if not cleaned or cleaned == "<hidden>" or "\ufffd" in raw_text:
        return -1
    control_count = sum(1 for ch in raw_text if unicodedata.category(ch).startswith("C"))
    if control_count > max(1, len(raw_text) // 4):
        return -1
    visible_ratio = len(cleaned) / max(1, len(raw_text))
    if visible_ratio < 0.55:
        return -1
    useful = sum(1 for ch in cleaned if ch.isalnum())
    if useful == 0:
        return -1
    return useful * 3 + len(cleaned)

def decode_ssid_bytes(data: bytes) -> str:
    if not data:
        return "<hidden>"
    best_text = ""
    best_score = -1
    for encoding in ("utf-8", "gb18030", "latin1"):
        try:
            raw_text = data.decode(encoding, errors="strict")
        except UnicodeDecodeError:
            continue
        cleaned = clean_ssid_text(raw_text)
        score = ssid_candidate_score(raw_text, cleaned)
        if score > best_score:
            best_text = cleaned
            best_score = score
    return best_text if best_score >= 0 else "<hidden>"

# 从管理帧的信息元素中解析 SSID。
def parse_ssid(mpdu: bytes, subtype: int) -> str:
    if len(mpdu) < 24:
        return ""
    if subtype in (8, 5):
        pos = 24 + 12
    elif subtype == 4:
        pos = 24
    elif subtype == 0:
        pos = 24 + 4
    elif subtype == 2:
        pos = 24 + 10
    else:
        return ""
    while pos + 2 <= len(mpdu):
        eid = mpdu[pos]
        elen = mpdu[pos + 1]
        pos += 2
        if pos + elen > len(mpdu):
            return ""
        if eid == 0 and elen <= 32:
            return "<hidden>" if elen == 0 else decode_ssid_bytes(mpdu[pos:pos + elen])
        pos += elen
    return ""

# 解析 802.11 MAC 帧头，提取地址字段、ToDS/FromDS、序列号、SSID 和节点关系。
def parse_mac_frame(mpdu: bytes) -> MacInfo:
    info = MacInfo()
    if len(mpdu) < 10:
        return info
    fc = mpdu[0] | (mpdu[1] << 8)
    if (fc & 3) != 0:
        return info
    typ = (fc >> 2) & 3
    sub = (fc >> 4) & 15
    to_ds = (fc >> 8) & 1
    from_ds = (fc >> 9) & 1
    info.type_num = typ
    info.subtype_num = sub
    info.type_name = type_name(typ)
    info.subtype_name = subtype_name(typ, sub)
    info.to_ds = to_ds
    info.from_ds = from_ds
    info.duration = mpdu[2] | (mpdu[3] << 8)
    info.addr1 = mac_at(mpdu, 4)

    if typ in (0, 2):
        if len(mpdu) < 24:
            return info
        info.addr2 = mac_at(mpdu, 10)
        info.addr3 = mac_at(mpdu, 16)
        if to_ds and from_ds:
            if len(mpdu) < 30:
                return info
            info.addr4 = mac_at(mpdu, 24)
        seq_ctl = mpdu[22] | (mpdu[23] << 8)
        info.fragment_number = seq_ctl & 15
        info.sequence_number = seq_ctl >> 4
    elif typ == 1:
        if len(mpdu) >= 16:
            info.addr2 = mac_at(mpdu, 10)
    else:
        return info

    if typ == 0:
        info.ra = info.addr1
        info.ta = info.addr2
        info.source_node = info.addr2
        info.destination_node = info.addr1
        info.bssid = info.addr3
        info.ssid = parse_ssid(mpdu, sub)
    elif typ == 2:
        info.ra = info.addr1
        info.ta = info.addr2
        if to_ds == 0 and from_ds == 0:
            info.destination_node = info.addr1
            info.source_node = info.addr2
            info.bssid = info.addr3
        elif to_ds == 1 and from_ds == 0:
            info.bssid = info.addr1
            info.source_node = info.addr2
            info.destination_node = info.addr3
        elif to_ds == 0 and from_ds == 1:
            info.destination_node = info.addr1
            info.bssid = info.addr2
            info.source_node = info.addr3
        else:
            info.destination_node = info.addr3
            info.source_node = info.addr4
            info.bssid = ""
    elif typ == 1:
        info.ra = info.addr1
        info.destination_node = info.addr1
        if info.addr2:
            info.ta = info.addr2
            info.source_node = info.addr2
    info.valid = True
    return info

def direction_text(src: str, dst: str) -> str:
    if src and dst:
        return f"{src} -> {dst}"
    if dst:
        return f"unknown -> {dst}"
    if src:
        return f"{src} -> unknown"
    return ""

# 将 MAC 解析结果和物理层解码参数合并成一条 Excel 结果记录。
def make_entry(info: MacInfo, pkt_start: int, mcs: int, length: int, bit_order: str, fcs_ok: bool, cfo: float) -> ResultEntry:
    return ResultEntry(
        ssid=info.ssid,
        bssid=info.bssid,
        frame_type=info.type_name,
        frame_subtype=info.subtype_name,
        source_node=info.source_node,
        destination_node=info.destination_node,
        direction=direction_text(info.source_node, info.destination_node),
        receiver_addr=info.ra,
        transmitter_addr=info.ta,
        addr1=info.addr1,
        addr2=info.addr2,
        addr3=info.addr3,
        addr4=info.addr4,
        to_ds=info.to_ds,
        from_ds=info.from_ds,
        fcs_ok=1 if fcs_ok else 0,
        start_sample=int(pkt_start),
        mcs=int(mcs),
        psdu_length=int(length),
        bit_order=bit_order,
        cfo_hz=float(cfo),
        sequence_number=info.sequence_number,
        fragment_number=info.fragment_number,
        duration=info.duration,
    )

# 生成结果去重用的键，避免同一个帧被不同候选重复保存。
def frame_key(entry: ResultEntry) -> str:
    return "|".join(map(str, [
        entry.start_sample, entry.frame_type, entry.frame_subtype,
        entry.source_node, entry.destination_node, entry.receiver_addr,
        entry.transmitter_addr, entry.bssid, entry.ssid,
    ]))

def near_duplicate_index(results: Sequence[ResultEntry], entry: ResultEntry, tolerance: int = 48) -> Optional[int]:
    for idx, old in enumerate(results):
        if abs(int(old.start_sample) - int(entry.start_sample)) > tolerance:
            continue
        if (
            old.frame_type == entry.frame_type
            and old.frame_subtype == entry.frame_subtype
            and old.source_node == entry.source_node
            and old.destination_node == entry.destination_node
            and old.receiver_addr == entry.receiver_addr
            and old.transmitter_addr == entry.transmitter_addr
            and old.bssid == entry.bssid
            and old.mcs == entry.mcs
            and old.psdu_length == entry.psdu_length
        ):
            return idx
    return None

def add_result_entry(seen: Dict[str, int], results: List[ResultEntry], entry: ResultEntry) -> bool:
    key = frame_key(entry)
    if key in seen:
        old = seen[key]
        if entry.fcs_ok and results[old].fcs_ok == 0:
            results[old] = entry
        return False
    near = near_duplicate_index(results, entry)
    if near is not None:
        if entry.fcs_ok and results[near].fcs_ok == 0:
            results[near] = entry
        return False
    seen[key] = len(results)
    results.append(entry)
    return True

# 在候选长度、MCS、bit 顺序和起点偏移中尝试解码，并把可信帧加入结果列表。
def try_decode_candidates(
    rx: np.ndarray,
    pkt_start: int,
    data_h: np.ndarray,
    pilot_h: np.ndarray,
    len_list: Sequence[int],
    mcs_list: Sequence[int],
    byte_orders: Sequence[bool],
    cfo_hz: float,
    opt: DecodeOptions,
    seen: Dict[str, int],
    results: List[ResultEntry],
    search_offsets: bool = False,
) -> bool:
    packet_decoded = False
    start_candidates = [pkt_start]
    if search_offsets:
        start_candidates += [pkt_start + delta for delta in range(-24, 25) if delta != 0]
    for length in len_list:
        if length <= 0 or length > 4095:
            continue
        for mcs in mcs_list:
            if mcs not in MCS_TABLE:
                continue
            control_rescue = opt.rescue_control_frames and ((mcs == 0 and length in (14, 20, 32)) or (mcs in (4, 5) and length == 20))
            cpe_modes = (0, 1, 2) if (control_rescue or opt.fallback_on_miss or search_offsets) else (0,)
            nsym = mcs_symbol_count(mcs, length)
            for data_start in start_candidates:
                if data_start < 0:
                    continue
                end = data_start + 400 + nsym * 80
                if end > rx.size:
                    continue
                pkt = apply_cfo(rx[data_start:end], cfo_hz, 0)
                cand_data_h, cand_pilot_h = estimate_channel(pkt[:400])
                for cpe_mode in cpe_modes:
                    decoded = recover_data_bits(pkt, cand_data_h, cand_pilot_h, mcs, length, cpe_mode=cpe_mode)
                    if decoded is None:
                        continue
                    saw_valid_unkept = False
                    pending_entries: List[ResultEntry] = []
                    for lsb_first in byte_orders:
                        bit_order = "lsb_first" if lsb_first else "msb_first"
                        mpdu = descramble_payload(decoded, length, lsb_first=lsb_first)
                        if not mpdu:
                            continue
                        info = parse_mac_frame(mpdu)
                        if not info.valid:
                            continue
                        saw_valid_unkept = True
                        fcs_ok = check_fcs(mpdu)
                        keep = fcs_ok or (info.type_num == 0 and info.subtype_num in (8, 5) and bool(info.ssid))
                        if not keep:
                            continue
                        entry = make_entry(info, data_start + 1, mcs, length, bit_order, fcs_ok, cfo_hz)
                        if (not fcs_ok) and opt.fallback_on_miss and not search_offsets:
                            pending_entries.append(entry)
                            continue
                        if not add_result_entry(seen, results, entry):
                            continue
                        packet_decoded = True
                        if opt.verbose_frames:
                            print(f"FRAME_FOUND start={entry.start_sample} mcs={mcs} len={length} "
                                  f"type={entry.frame_type}/{entry.frame_subtype} "
                                  f"source={entry.source_node} destination={entry.destination_node} "
                                  f"bssid={entry.bssid} ssid={entry.ssid} fcs={entry.fcs_ok}")
                        if opt.stop_after_fcs_ok and fcs_ok:
                            break
                    if packet_decoded and opt.stop_after_fcs_ok:
                        break
                    if saw_valid_unkept and opt.fallback_on_miss and not search_offsets:
                        if try_decode_candidates(
                            rx, pkt_start, data_h, pilot_h, [length], [mcs],
                            byte_orders, cfo_hz, opt, seen, results, search_offsets=True,
                        ):
                            return True
                    for entry in pending_entries:
                        if add_result_entry(seen, results, entry):
                            packet_decoded = True
                            if opt.verbose_frames:
                                print(f"FRAME_FOUND start={entry.start_sample} mcs={mcs} len={length} "
                                      f"type={entry.frame_type}/{entry.frame_subtype} "
                                      f"source={entry.source_node} destination={entry.destination_node} "
                                      f"bssid={entry.bssid} ssid={entry.ssid} fcs={entry.fcs_ok}")
                    if packet_decoded and opt.stop_after_fcs_ok:
                        break
                if packet_decoded and opt.stop_after_fcs_ok:
                    break
            if packet_decoded and opt.fast_mode:
                break
        if packet_decoded and opt.fast_mode:
            break
    return packet_decoded

def print_progress(pos: int, total: int, packets: int) -> None:
    pct = 100.0 * pos / max(total, 1)
    print(f"解析进度: 当前索引={pos} / 全信号长度={total}, 进度={pct:.2f}%, 已检查包数={packets}")

# 生成 0、-1、+1、-2、+2 这样的近邻偏移序列，用于提示起点附近微调。
def nearby_deltas(radius: int) -> List[int]:
    values = [0]
    for step in range(1, radius + 1):
        values.extend([-step, step])
    return values

# 使用已有 output.xlsx 中的起始采样点等提示进行快速解码。
def decode_known_starts(capture_file: Path, opt: DecodeOptions, hints: Sequence[HintEntry]) -> List[ResultEntry]:
    rx = read_capture(capture_file, "sc16", opt.max_complex_samples)
    print(f"complex_samples_20msps: {rx.size}")
    print(f"hint_starts: {len(hints)}")
    results: List[ResultEntry] = []
    seen: Dict[str, int] = {}
    packets = 0
    lsig_ok = 0
    max_packets = opt.max_packets if opt.max_packets is not None else 10**12
    max_frames = opt.max_frames if opt.max_frames is not None else 10**12
    for idx, hint in enumerate(hints, 1):
        if packets >= max_packets or len(results) >= max_frames:
            break
        if idx == 1 or idx % 20 == 0:
            print(f"hint解析进度: {idx}/{len(hints)}, 已检查包数={packets}")
        start1 = int(hint.start_sample)
        base = start1 - 1
        best_entry: Optional[ResultEntry] = None
        local_opt = DecodeOptions(**asdict(opt))
        local_opt.fallback_on_miss = False
        for delta in nearby_deltas(24):
            pkt_start = base + delta
            if pkt_start < 0 or pkt_start + 400 > rx.size:
                continue
            try:
                cfo, pre_cf, data_h, pilot_h = estimate_frontend(rx, pkt_start)
                if hint.psdu_length is not None and hint.mcs is not None:
                    len_list = filter_lengths([hint.psdu_length], opt)
                    mcs_list = [hint.mcs] if hint.mcs in MCS_TABLE else []
                else:
                    lsig_bits = recover_lsig_bits(pre_cf, data_h, pilot_h)
                    if lsig_bits is None:
                        continue
                    len_list = filter_lengths(candidate_lengths(lsig_bits, False), opt)
                    mcs_list = candidate_mcs(lsig_bits, opt.exhaustive_mcs)
                if not len_list or not mcs_list:
                    continue
                lsig_ok += 1
                temp_results: List[ResultEntry] = []
                temp_seen: Dict[str, int] = {}
                try_decode_candidates(
                    rx, pkt_start, data_h, pilot_h, len_list, mcs_list,
                    [True, False] if opt.try_both_byte_orders else [True],
                    cfo, local_opt, temp_seen, temp_results, search_offsets=False,
                )
                if temp_results:
                    candidate = sorted(temp_results, key=lambda r: (r.fcs_ok, -abs(r.start_sample - int(start1))), reverse=True)[0]
                    if best_entry is None or (candidate.fcs_ok, -abs(candidate.start_sample - int(start1))) > (best_entry.fcs_ok, -abs(best_entry.start_sample - int(start1))):
                        best_entry = candidate
                if best_entry is not None and best_entry.fcs_ok:
                    break
            except Exception:
                continue
        if best_entry is not None:
            packets += 1
            best_entry.start_sample = int(start1)
            key = frame_key(best_entry)
            if key in seen:
                old = seen[key]
                if best_entry.fcs_ok and results[old].fcs_ok == 0:
                    results[old] = best_entry
            else:
                seen[key] = len(results)
                results.append(best_entry)
    print(
        f"packets_checked={packets} lsig_ok={lsig_ok} "
        f"frames_saved={len(results)} fcs_ok={sum(r.fcs_ok for r in results)}"
    )
    return results

# 不依赖提示表，直接扫描完整 IQ 信号并尝试解码 WiFi 包。
def decode_capture(capture_file: Path, opt: DecodeOptions) -> List[ResultEntry]:
    rx = read_capture(capture_file, "sc16", opt.max_complex_samples)
    print(f"complex_samples_20msps: {rx.size}")
    results: List[ResultEntry] = []
    seen: Dict[str, int] = {}
    pos = 0
    packets = 0
    lsig_ok = 0
    total = rx.size
    max_packets = opt.max_packets if opt.max_packets is not None else 10**12
    last_report = 0.0
    noise_floor = max(float(np.median(np.abs(rx[::max(1, rx.size // 120_000)]) ** 2)), 1e-8)
    print_progress(1, total, packets)

    max_frames = opt.max_frames if opt.max_frames is not None else 10**12
    while pos < total - 6000 and packets < max_packets and len(results) < max_frames:
        pct = 100.0 * pos / total
        if pct >= last_report + opt.progress_step_percent:
            print_progress(pos + 1, total, packets)
            last_report = pct
        detect_end = min(total, pos + opt.detect_window_samples)
        window = rx[pos:detect_end]
        if not window_has_signal(window, noise_floor, opt.energy_gate_factor):
            if detect_end >= total:
                break
            pos = max(pos + 1, detect_end - opt.detect_overlap_samples)
            continue
        offsets = lltf_candidate_offsets(window, max_candidates=opt.lltf_max_candidates, earliest_only=True)
        if not offsets:
            offsets = detect_packet_offsets(window, max_candidates=opt.stf_max_candidates)
        if not offsets:
            if detect_end >= total:
                break
            pos = max(pos + 1, detect_end - opt.detect_overlap_samples)
            continue
        advanced = False
        last_candidate_start = pos + offsets[-1]
        for offset in offsets:
            if packets >= max_packets or len(results) >= max_frames:
                break
            pkt_start = refine_packet_start(rx, pos + offset)
            if pkt_start + 400 > total:
                break

            try:
                cfo, pre_cf, data_h, pilot_h = estimate_frontend(rx, pkt_start)
                lsig_bits = recover_lsig_bits(pre_cf, data_h, pilot_h)
                if lsig_bits is None:
                    continue
                lsig_ok += 1
                packets += 1
                len_list = filter_lengths(candidate_lengths(lsig_bits, opt.fast_mode), opt)
                mcs_list = candidate_mcs(lsig_bits, opt.exhaustive_mcs)
                if not len_list or not mcs_list:
                    continue
                byte_orders = [True, False] if opt.try_both_byte_orders else [True]
                decoded = try_decode_candidates(
                    rx, pkt_start, data_h, pilot_h, len_list, mcs_list,
                    byte_orders, cfo, opt, seen, results,
                )
                if (not decoded) and opt.fallback_on_miss:
                    len_list = filter_lengths(candidate_lengths(lsig_bits, False), opt)
                    mcs_list = sorted(set(mcs_list) | set(range(8)))
                    try_decode_candidates(
                        rx, pkt_start, data_h, pilot_h, len_list, mcs_list,
                        [True, False], cfo, opt, seen, results,
                    )
                # 和 MATLAB 版本保持一致：每处理一个候选包后只前进约一个前导长度。
                # 如果按 L-SIG 估计出的整包长度大步跳，忙信道里紧随其后的 ACK/CTS/RTS
                # 很容易被直接跨过，这正是 Python 旧版比 MATLAB 少帧的主要原因之一。
                pos = min(total, pkt_start + max(80, int(opt.scan_step_samples)))
                advanced = True
                break
            except Exception:
                continue
        if not advanced:
            pos = max(pos + 1, min(detect_end - opt.detect_overlap_samples, last_candidate_start + 160))

    print_progress(min(pos + 1, total), total, packets)
    print(
        f"packets_checked={packets} lsig_ok={lsig_ok} "
        f"frames_saved={len(results)} fcs_ok={sum(r.fcs_ok for r in results)}"
    )
    return results

def yes_no(value: int) -> str:
    return "是" if value else "否"

# 把 ResultEntry 列表整理成 output.xlsx 需要的二维表格数据。
def result_rows(results: List[ResultEntry], capture_file: Path) -> List[List[object]]:
    if not results:
        row = [""] * len(HEADERS)
        row[1] = "未解析到可信WiFi MAC帧"
        row[-1] = str(capture_file)
        return [row]
    rows: List[List[object]] = []
    for idx, r in enumerate(results, 1):
        rows.append([
            idx, r.ssid, r.bssid, r.frame_type, r.frame_subtype,
            r.source_node, r.destination_node, r.direction, r.receiver_addr, r.transmitter_addr,
            r.addr1, r.addr2, r.addr3, r.addr4, r.to_ds, r.from_ds,
            yes_no(r.fcs_ok), r.start_sample, r.mcs, r.psdu_length, r.bit_order,
            r.cfo_hz, r.sequence_number if r.sequence_number is not None else "",
            r.fragment_number if r.fragment_number is not None else "",
            r.duration if r.duration is not None else "", str(capture_file),
        ])
    return rows

# 把多文件解码结果整理成同一个 output.xlsx；每一行保留对应的采集文件路径。
def result_rows_multi(capture_results: Sequence[CaptureResult], empty_captures: Sequence[Path]) -> List[List[object]]:
    rows: List[List[object]] = []
    for idx, item in enumerate(capture_results, 1):
        r = item.result
        rows.append([
            idx, r.ssid, r.bssid, r.frame_type, r.frame_subtype,
            r.source_node, r.destination_node, r.direction, r.receiver_addr, r.transmitter_addr,
            r.addr1, r.addr2, r.addr3, r.addr4, r.to_ds, r.from_ds,
            yes_no(r.fcs_ok), r.start_sample, r.mcs, r.psdu_length, r.bit_order,
            r.cfo_hz, r.sequence_number if r.sequence_number is not None else "",
            r.fragment_number if r.fragment_number is not None else "",
            r.duration if r.duration is not None else "", str(item.capture_file),
        ])
    for capture_file in empty_captures:
        row = [""] * len(HEADERS)
        row[0] = len(rows) + 1
        row[1] = "未解析到可信WiFi MAC帧"
        row[-1] = str(capture_file)
        rows.append(row)
    return rows

def safe_node_id(value: str) -> str:
    return value.strip().lower()

def join_unique(values: Sequence[str]) -> str:
    cleaned_values: List[str] = []
    saw_hidden = False
    for value in values:
        cleaned = clean_ssid_text(value)
        if cleaned == "<hidden>":
            saw_hidden = True
        elif cleaned:
            cleaned_values.append(cleaned)
    unique = sorted(set(cleaned_values))
    if unique:
        if len(unique) > 4:
            unique = unique[:4] + [f"+{len(unique) - 4} more"]
        return ", ".join(unique)
    return "<hidden>" if saw_hidden else ""

def node_kind(node_id: str, ssids: Sequence[str], bssids: Sequence[str]) -> str:
    if node_id == "ff:ff:ff:ff:ff:ff":
        return "Broadcast"
    if node_id in {b.lower() for b in bssids if b}:
        cleaned_ssids = [clean_ssid_text(ssid) for ssid in ssids]
        cleaned_ssids = [ssid for ssid in cleaned_ssids if ssid]
        return "Hidden AP" if cleaned_ssids and set(cleaned_ssids) == {"<hidden>"} else "Access Point"
    return "Station"

def mac_oui_prefix(node_id: str) -> str:
    parts = safe_node_id(node_id).split(":")
    return ":".join(parts[:3]) if len(parts) >= 3 else ""

def is_local_admin_mac(node_id: str) -> bool:
    parts = safe_node_id(node_id).split(":")
    if not parts:
        return False
    try:
        return bool(int(parts[0], 16) & 0x02)
    except ValueError:
        return False

def is_intel_mac(node_id: str) -> bool:
    return (not is_local_admin_mac(node_id)) and mac_oui_prefix(node_id) in INTEL_OUI_PREFIXES

def node_vendor(node_id: str, kind: str) -> str:
    return "Intel" if kind == "Station" and is_intel_mac(node_id) else ""

def node_device_icon(node_id: str, kind: str) -> str:
    return "intel_laptop" if node_vendor(node_id, kind) == "Intel" else ""

def plot_stem(node_id: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", safe_node_id(node_id)).strip("_") or "node"

def capture_plot_stem(capture_file: Path) -> str:
    return re.sub(r"[^a-z0-9]+", "_", capture_file.stem.lower()).strip("_") or "capture"

def capture_channel_meta(capture_file: Path) -> Dict[str, object]:
    match = re.search(
        r"wifi_(?P<band>[^_]+)_ch(?P<channel>\d+)_(?P<freq>\d+)MHz_(?P<rate>\d+)Msps_(?P<duration>[\d.]+)s",
        capture_file.stem,
        flags=re.IGNORECASE,
    )
    if not match:
        return {
            "label": capture_file.stem,
            "band": "",
            "channel": "",
            "freq_mhz": "",
            "sample_rate_msps": "",
            "duration_s": "",
        }
    band_raw = match.group("band").lower()
    band = "2.4G" if band_raw in {"2g4", "2.4g"} else ("5G" if band_raw == "5g" else band_raw.upper())
    channel = int(match.group("channel"))
    freq_mhz = int(match.group("freq"))
    rate_msps = int(match.group("rate"))
    duration_s = float(match.group("duration"))
    return {
        "label": f"{band} CH{channel:03d} / {freq_mhz} MHz",
        "band": band,
        "channel": channel,
        "freq_mhz": freq_mhz,
        "sample_rate_msps": rate_msps,
        "duration_s": duration_s,
    }

def icon_source_candidates(icon_name: str) -> List[Path]:
    script_dir = Path(__file__).resolve().parent
    return [
        script_dir / "assets" / icon_name,
        script_dir / icon_name,
        script_dir.parent / "every wifi node signal" / "tp" / icon_name,
    ]

def copy_asset_icon(output_dir: Path, icon_name: str) -> None:
    target_dir = output_dir / "assets"
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / icon_name
    for source in icon_source_candidates(icon_name):
        if source.exists():
            if source.resolve() != target.resolve():
                target.write_bytes(source.read_bytes())
            return

def copy_topology_icons(output_dir: Path) -> None:
    copy_asset_icon(output_dir, ROUTER_ICON_NAME)
    copy_asset_icon(output_dir, LAPTOP_ICON_NAME)

# 估算一个 WiFi 帧在原始采样中的长度，供节点信号切片和绘图使用。
def frame_sample_count(entry: ResultEntry) -> int:
    if entry.mcs in MCS_TABLE and entry.psdu_length > 0:
        return 400 + mcs_symbol_count(entry.mcs, entry.psdu_length) * 80
    return 1600

# 找出一个解码帧和哪些节点有关：源、目的、BSSID、RA、TA 都算作相关节点。
def related_node_ids(entry: ResultEntry) -> List[str]:
    values = [entry.source_node, entry.destination_node, entry.bssid, entry.receiver_addr, entry.transmitter_addr]
    nodes: List[str] = []
    seen = set()
    for value in values:
        node_id = safe_node_id(value)
        if node_id and node_id not in seen:
            seen.add(node_id)
            nodes.append(node_id)
    return nodes

def frame_window(entry: ResultEntry, rx_size: int, margin: float = 0.0) -> Tuple[int, int]:
    start = max(0, int(entry.start_sample))
    length = max(1, frame_sample_count(entry))
    extra = int(round(length * margin))
    left = max(0, start - extra)
    right = min(rx_size, start + length + extra)
    return left, max(left + 1, right)

def group_results_by_node(results: Sequence[ResultEntry]) -> Dict[str, List[ResultEntry]]:
    grouped: Dict[str, List[ResultEntry]] = {}
    for entry in results:
        for node_id in related_node_ids(entry):
            grouped.setdefault(node_id, []).append(entry)
    return grouped

# 生成某个节点在整段采集中的“大尺度”抽取信号：非该节点相关帧的位置保持为 0。
def node_timeline_envelope(rx_abs: np.ndarray, entries: Sequence[ResultEntry], bins: int) -> Tuple[np.ndarray, np.ndarray]:
    total = int(rx_abs.size)
    bins = max(1000, min(int(bins), max(total, 1)))
    edges = np.linspace(0, total, bins + 1, dtype=np.int64)
    envelope = np.zeros(bins, dtype=np.float32)
    for entry in entries:
        left, right = frame_window(entry, total)
        first = max(0, min(bins - 1, left * bins // max(total, 1)))
        last = max(first + 1, min(bins, (right * bins + total - 1) // max(total, 1)))
        for idx in range(first, last):
            a = max(left, int(edges[idx]))
            b = min(right, int(edges[idx + 1]))
            if b > a:
                envelope[idx] = max(envelope[idx], float(np.max(rx_abs[a:b])))
    times = (edges[:-1] + edges[1:]) * 0.5 / SAMPLE_RATE
    return times, envelope

# 给每个拓扑节点写出两张 PNG：整段信号下的节点抽取图，以及一帧完整波形的特写图。
def generate_node_signal_plots(
    results: Sequence[ResultEntry],
    capture_file: Path,
    output_dir: Path,
    max_complex: Optional[int] = None,
) -> Dict[str, Dict[str, str]]:
    if not results:
        return {}
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except Exception as exc:
        print(f"Signal plot generation skipped: matplotlib is not available ({exc})")
        return {}

    rx = read_capture(capture_file, "sc16", max_complex)
    rx_abs = np.abs(rx).astype(np.float32)
    grouped = group_results_by_node(results)
    plot_dir = output_dir / "signal_plots"
    plot_dir.mkdir(parents=True, exist_ok=True)
    plot_paths: Dict[str, Dict[str, str]] = {}
    timeline_bins = min(12000, max(3000, int(rx.size // 2500) or 3000))

    for node_id, entries in grouped.items():
        stem = plot_stem(node_id)
        timeline_file = plot_dir / f"{stem}_timeline.png"
        frame_file = plot_dir / f"{stem}_frame.png"

        times, envelope = node_timeline_envelope(rx_abs, entries, timeline_bins)
        fig, ax = plt.subplots(figsize=(10.5, 3.2), dpi=150)
        ax.plot(times, envelope, color="#2563eb", linewidth=1.05)
        ax.fill_between(times, 0, envelope, color="#93c5fd", alpha=0.34, linewidth=0)
        ax.set_title(f"Node {node_id} extracted signal in full capture", fontsize=10)
        ax.set_xlabel("Time in full capture (s)")
        ax.set_ylabel("Normalized amplitude")
        ax.set_xlim(0, rx.size / SAMPLE_RATE)
        ax.set_ylim(0, max(1e-6, float(np.max(envelope)) * 1.12))
        ax.grid(True, color="#d8e0ea", linewidth=0.6, alpha=0.85)
        fig.tight_layout()
        fig.savefig(timeline_file)
        plt.close(fig)

        frame_entry = min(entries, key=lambda item: int(item.start_sample))
        left, right = frame_window(frame_entry, rx.size, margin=0.10)
        segment = rx[left:right]
        t_us = (np.arange(left, right) - int(frame_entry.start_sample)) / SAMPLE_RATE * 1e6
        fig, ax = plt.subplots(figsize=(10.5, 3.2), dpi=150)
        ax.plot(t_us, np.abs(segment), color="#079669", linewidth=1.05)
        ax.axvline(0, color="#111827", linewidth=0.8, alpha=0.55)
        ax.set_title(f"Node {node_id} single frame detail", fontsize=10)
        ax.set_xlabel("Time relative to frame start (us)")
        ax.set_ylabel("Normalized |IQ|")
        ax.grid(True, color="#d8e0ea", linewidth=0.6, alpha=0.85)
        fig.tight_layout()
        fig.savefig(frame_file)
        plt.close(fig)

        plot_paths[node_id] = {
            "timeline_plot": f"signal_plots/{timeline_file.name}",
            "frame_plot": f"signal_plots/{frame_file.name}",
            "signal_plot": f"signal_plots/{timeline_file.name}",
        }
    return plot_paths

# 多文件场景下给节点生成信号图：同一个节点可能出现在多个信道里，这里选它第一次出现的文件来画图。
def generate_node_signal_plots_multi(
    capture_results: Sequence[CaptureResult],
    output_dir: Path,
    max_complex: Optional[int] = None,
) -> Dict[str, Dict[str, str]]:
    by_capture: Dict[Path, List[ResultEntry]] = {}
    for item in capture_results:
        by_capture.setdefault(item.capture_file, []).append(item.result)
    merged: Dict[str, Dict[str, str]] = {}
    for capture_file, results in by_capture.items():
        paths = generate_node_signal_plots(results, capture_file, output_dir, max_complex)
        for node_id, plot_path in paths.items():
            merged.setdefault(node_id, plot_path)
    return merged

# 给每个 WiFi 子信道采集文件生成一张整段 1 秒大尺度 |IQ| 时域包络图，供设置弹窗底部展示。
def generate_channel_waveform_plots(
    capture_files: Sequence[Path],
    output_dir: Path,
    max_complex: Optional[int] = None,
    bins: int = 9000,
) -> List[Dict[str, object]]:
    if not capture_files:
        return []
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except Exception as exc:
        print(f"Channel waveform plot generation skipped: matplotlib is not available ({exc})")
        return []

    plot_dir = output_dir / "signal_plots" / "channel_waveforms"
    plot_dir.mkdir(parents=True, exist_ok=True)
    items: List[Dict[str, object]] = []
    for idx, capture_file in enumerate(capture_files, 1):
        try:
            rx = read_capture(capture_file, "sc16", max_complex)
        except Exception as exc:
            print(f"Channel waveform plot skipped for {capture_file.name}: {exc}")
            continue
        rx_abs = np.abs(rx).astype(np.float32)
        total = int(rx_abs.size)
        if total <= 0:
            continue
        plot_bins = max(1000, min(int(bins), total))
        edges = np.linspace(0, total, plot_bins + 1, dtype=np.int64)
        peak = np.zeros(plot_bins, dtype=np.float32)
        mean = np.zeros(plot_bins, dtype=np.float32)
        for i in range(plot_bins):
            a = int(edges[i])
            b = int(edges[i + 1])
            if b <= a:
                continue
            segment = rx_abs[a:b]
            peak[i] = float(np.max(segment))
            mean[i] = float(np.mean(segment))
        times = (edges[:-1] + edges[1:]) * 0.5 / SAMPLE_RATE
        meta = capture_channel_meta(capture_file)
        image_file = plot_dir / f"{idx:02d}_{capture_plot_stem(capture_file)}_waveform.png"
        fig, ax = plt.subplots(figsize=(11.0, 3.0), dpi=145)
        ax.fill_between(times, 0, peak, color="#93c5fd", alpha=0.36, linewidth=0)
        ax.plot(times, peak, color="#2563eb", linewidth=0.78, label="Peak |IQ|")
        ax.plot(times, mean, color="#079669", linewidth=0.7, alpha=0.82, label="Mean |IQ|")
        ax.set_title(f"{meta['label']} full-capture time-domain waveform", fontsize=10)
        ax.set_xlabel("Time (s)")
        ax.set_ylabel("Normalized |IQ|")
        ax.set_xlim(0, total / SAMPLE_RATE)
        ax.set_ylim(0, max(1e-6, float(np.max(peak)) * 1.12))
        ax.grid(True, color="#d8e0ea", linewidth=0.55, alpha=0.85)
        ax.legend(loc="upper right", fontsize=7, frameon=True)
        fig.tight_layout()
        fig.savefig(image_file)
        plt.close(fig)
        items.append({
            **meta,
            "file": capture_file.name,
            "path": f"signal_plots/channel_waveforms/{image_file.name}",
            "samples": total,
        })
    return items

# 根据解码出的帧结果生成拓扑图数据：节点是 MAC 地址，边是源节点 -> 目的节点。
def build_topology(
    results: List[ResultEntry],
    capture_file: Path,
    plot_paths: Optional[Dict[str, Dict[str, str]]] = None,
    channel_plots: Optional[List[Dict[str, object]]] = None,
) -> Dict[str, object]:
    return build_topology_multi(
        [CaptureResult(capture_file, result) for result in results],
        [capture_file],
        plot_paths,
        capture_file.parent,
        channel_plots,
    )

# 根据多个采集文件的解码结果生成拓扑图数据：节点是 MAC 地址，边是源节点 -> 目的节点。
def build_topology_multi(
    capture_results: Sequence[CaptureResult],
    capture_files: Sequence[Path],
    plot_paths: Optional[Dict[str, Dict[str, str]]] = None,
    source_dir: Optional[Path] = None,
    channel_plots: Optional[List[Dict[str, object]]] = None,
) -> Dict[str, object]:
    node_stats: Dict[str, Dict[str, object]] = {}
    edge_stats: Dict[Tuple[str, str], Dict[str, object]] = {}
    frames: List[Dict[str, object]] = []

    def ensure_node(node_id: str) -> Dict[str, object]:
        node_id = safe_node_id(node_id)
        if not node_id:
            return {}
        if node_id not in node_stats:
            node_stats[node_id] = {
                "id": node_id, "label": "Broadcast" if node_id == "ff:ff:ff:ff:ff:ff" else node_id,
                "sent": 0, "received": 0, "frames": 0, "ssids": [], "bssids": [],
                "capture_files": set(), "start_samples": [],
            }
        return node_stats[node_id]

    for idx, item in enumerate(capture_results, 1):
        r = item.result
        capture_file = item.capture_file
        src = safe_node_id(r.source_node)
        dst = safe_node_id(r.destination_node)
        related = [src, dst, safe_node_id(r.bssid), safe_node_id(r.receiver_addr), safe_node_id(r.transmitter_addr)]
        for node_id in related:
            node = ensure_node(node_id)
            if not node:
                continue
            node["frames"] = int(node["frames"]) + 1
            node["capture_files"].add(str(capture_file))
            node["start_samples"].append(int(r.start_sample))
            if r.ssid:
                node["ssids"].append(r.ssid)
            if r.bssid:
                node["bssids"].append(safe_node_id(r.bssid))
        if src:
            ensure_node(src)["sent"] += 1
        if dst:
            ensure_node(dst)["received"] += 1
        if src and dst:
            key = (src, dst)
            edge = edge_stats.setdefault(key, {
                "id": f"{src}->{dst}", "source": src, "target": dst, "weight": 0,
                "frames": [], "ssids": [], "relation": f"{src} -> {dst}",
            })
            edge["weight"] += 1
            edge["frames"].append(idx)
            if r.ssid:
                edge["ssids"].append(r.ssid)
        frames.append({
            "index": idx, "ssid": r.ssid, "bssid": r.bssid, "frame_type": r.frame_type,
            "frame_subtype": r.frame_subtype, "source": src, "target": dst,
            "direction": r.direction, "receiver": r.receiver_addr, "transmitter": r.transmitter_addr,
            "fcs_ok": yes_no(r.fcs_ok), "start_sample": r.start_sample, "mcs": r.mcs,
            "psdu_length": r.psdu_length, "cfo_hz": r.cfo_hz,
        })

    nodes = []
    for node in node_stats.values():
        ssid = join_unique(node["ssids"])
        bssid = join_unique(node["bssids"])
        sent = int(node["sent"])
        received = int(node["received"])
        target_only = sent == 0 and received > 0
        kind = node_kind(node["id"], node["ssids"], node["bssids"])
        item = {
            "id": node["id"], "label": node["label"], "kind": kind,
            "ssid": ssid, "bssid": bssid, "sent": sent, "received": received,
            "frames": node["frames"], "capture_files": sorted(node["capture_files"]),
            "start_samples": sorted(set(node["start_samples"]))[:80],
            "has_radiated_signal": sent > 0,
            "target_only": target_only,
            "vendor": node_vendor(node["id"], kind),
            "device_icon": node_device_icon(node["id"], kind),
        }
        if plot_paths and node["id"] in plot_paths:
            item.update(plot_paths[node["id"]])
        nodes.append(item)
    edges = [{
        "id": e["id"], "source": e["source"], "target": e["target"], "weight": e["weight"],
        "label": str(e["weight"]), "relation": e["relation"], "ssid": join_unique(e["ssids"]),
        "frames": e["frames"],
    } for e in edge_stats.values()]
    nodes.sort(key=lambda n: (n["kind"] != "Access Point", n["kind"] != "Hidden AP", n["id"]))
    edges.sort(key=lambda e: (-e["weight"], e["source"], e["target"]))
    return {
        "source_xlsx": "output.xlsx", "source_dir": str(source_dir or (capture_files[0].parent if capture_files else "")),
        "nodes": nodes, "edges": edges, "frames": frames, "channel_plots": channel_plots or [],
        "summary": {
            "node_count": len(nodes), "edge_count": len(edges), "frame_count": len(frames),
            "capture_files": [str(path) for path in capture_files],
        },
    }

def topology_html_legacy(graph: Dict[str, object]) -> str:
    graph_json = json.dumps(graph, ensure_ascii=False)
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>WiFi Signal Topology</title>
<style>
body{{margin:0;font-family:Segoe UI,Arial,sans-serif;background:#f7f8fb;color:#172033}}
.app{{display:grid;grid-template-columns:280px 1fr 360px;height:100vh}}
aside,.details{{background:#fff;border-right:1px solid #d9dee8;overflow:auto}}
.details{{border-right:0;border-left:1px solid #d9dee8}}
h1{{font-size:17px;margin:18px 16px 8px}} .stats{{display:flex;gap:8px;flex-wrap:wrap;margin:0 16px 14px}}
.stats span{{background:#eef3ff;border:1px solid #cbd8ff;border-radius:6px;padding:5px 8px;font-size:12px}}
.search{{margin:0 16px 12px;width:calc(100% - 32px);box-sizing:border-box;padding:8px;border:1px solid #cfd6e3;border-radius:6px}}
.item{{padding:10px 16px;border-top:1px solid #edf0f5;cursor:pointer;font-size:13px}}
.item:hover,.item.active{{background:#eef6ff}} .muted{{color:#667085;font-size:12px}}
main{{position:relative;overflow:hidden}} svg{{width:100%;height:100%;background:#fbfcff}}
.edge{{stroke:#98a2b3;stroke-opacity:.72}} .edge.selected{{stroke:#dc2626;stroke-opacity:1}}
.node{{stroke:#fff;stroke-width:2;cursor:pointer;filter:drop-shadow(0 2px 4px #0002)}} .node.selected{{stroke:#111827;stroke-width:3}}
.label{{font-size:11px;text-anchor:middle;fill:#1f2937;pointer-events:none}}
.edgeLabel{{font-size:11px;fill:#475467;pointer-events:none}}
.panel{{padding:16px}} table{{width:100%;border-collapse:collapse;font-size:12px}} td{{border-bottom:1px solid #edf0f5;padding:7px 4px;vertical-align:top}}
td:first-child{{color:#667085;width:105px}} .empty{{padding:16px;color:#667085;font-size:13px}}
.frame{{font-size:12px;border-bottom:1px solid #edf0f5;padding:8px 0}}
</style>
</head>
<body>
<div class="app">
<aside>
  <h1>WiFi Topology</h1>
  <div class="stats" id="stats"></div>
  <input class="search" id="filter" placeholder="Filter MAC / SSID">
  <div id="nodeList"></div>
</aside>
<main><svg id="graph"></svg></main>
<section class="details"><div id="details" class="empty">点击节点或连线查看详情。</div></section>
</div>
<script id="graph-data" type="application/json">{graph_json}</script>
<script>
const graph = JSON.parse(document.getElementById("graph-data").textContent);
const svg = document.getElementById("graph"), details = document.getElementById("details");
const stats = document.getElementById("stats"), nodeList = document.getElementById("nodeList"), filter = document.getElementById("filter");
let selected = null;
const nodeMap = new Map(graph.nodes.map(n => [n.id, n]));
const edgeMap = new Map(graph.edges.map(e => [e.id, e]));
stats.innerHTML = `<span>${{graph.summary.node_count}} nodes</span><span>${{graph.summary.edge_count}} edges</span><span>${{graph.summary.frame_count}} frames</span>`;
function esc(v){{return String(v ?? "").replace(/[&<>"]/g,c=>({{"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;"}}[c]));}}
function color(n){{return n.kind==="Access Point"?"#2563eb":n.kind==="Hidden AP"?"#7c3aed":n.kind==="Broadcast"?"#ef4444":"#059669";}}
function radius(n){{return Math.max(13, Math.min(34, 11 + Math.sqrt(n.frames || 1) * 3));}}
function layout(){{
  const w = svg.clientWidth || 900, h = svg.clientHeight || 650, cx=w/2, cy=h/2;
  const aps = graph.nodes.filter(n=>n.kind==="Access Point"||n.kind==="Hidden AP");
  const others = graph.nodes.filter(n=>!aps.includes(n)&&n.kind!=="Broadcast");
  const bcasts = graph.nodes.filter(n=>n.kind==="Broadcast");
  const pos = new Map();
  aps.forEach((n,i)=>{{const a=2*Math.PI*i/Math.max(aps.length,1)-Math.PI/2; pos.set(n.id,{{x:cx+150*Math.cos(a),y:cy+115*Math.sin(a)}});}});
  others.forEach((n,i)=>{{const a=2*Math.PI*i/Math.max(others.length,1)-Math.PI/2; pos.set(n.id,{{x:cx+Math.min(w,h)*0.38*Math.cos(a),y:cy+Math.min(w,h)*0.34*Math.sin(a)}});}});
  bcasts.forEach((n,i)=>pos.set(n.id,{{x:w-110,y:90+i*75}}));
  return pos;
}}
function draw(){{
  const pos = layout(); svg.innerHTML = "";
  for (const e of graph.edges) {{
    const a=pos.get(e.source), b=pos.get(e.target); if(!a||!b) continue;
    const line=document.createElementNS("http://www.w3.org/2000/svg","line");
    line.setAttribute("x1",a.x); line.setAttribute("y1",a.y); line.setAttribute("x2",b.x); line.setAttribute("y2",b.y);
    line.setAttribute("stroke-width",Math.max(1.4,Math.min(8,1+Math.sqrt(e.weight))));
    line.setAttribute("class","edge"+(selected===e.id?" selected":"")); line.onclick=()=>selectEdge(e.id); svg.appendChild(line);
    if(e.weight>1){{const t=document.createElementNS("http://www.w3.org/2000/svg","text"); t.setAttribute("x",(a.x+b.x)/2); t.setAttribute("y",(a.y+b.y)/2-5); t.setAttribute("class","edgeLabel"); t.textContent=e.weight; svg.appendChild(t);}}
  }}
  for (const n of graph.nodes) {{
    const p=pos.get(n.id), c=document.createElementNS("http://www.w3.org/2000/svg","circle"); if(!p) continue;
    c.setAttribute("cx",p.x); c.setAttribute("cy",p.y); c.setAttribute("r",radius(n)); c.setAttribute("fill",color(n));
    c.setAttribute("class","node"+(selected===n.id?" selected":"")); c.onclick=()=>selectNode(n.id); svg.appendChild(c);
    const t=document.createElementNS("http://www.w3.org/2000/svg","text"); t.setAttribute("x",p.x); t.setAttribute("y",p.y+radius(n)+15); t.setAttribute("class","label"); t.textContent=n.label==="Broadcast"?"Broadcast":n.id.slice(-8); svg.appendChild(t);
  }}
}}
function rows(items){{return `<table>${{items.map(([k,v])=>`<tr><td>${{esc(k)}}</td><td>${{esc(v)}}</td></tr>`).join("")}}</table>`;}}
function nodeFrames(id){{return graph.frames.filter(f=>[f.source,f.target,f.bssid,f.receiver,f.transmitter].includes(id)).slice(0,18);}}
function selectNode(id){{
  selected=id; const n=nodeMap.get(id);
  details.innerHTML=`<div class="panel"><h2 style="font-size:16px;margin:0 0 10px">${{esc(n.label)}}</h2>`+
    rows([["Type",n.kind],["MAC",n.id],["SSID",n.ssid],["BSSID",n.bssid],["Sent",n.sent],["Received",n.received],["Frames",n.frames],["Capture",n.capture_files.join("\\n")],["Start samples",n.start_samples.join(", ")]])+
    `<h3>Related frames</h3>${{nodeFrames(id).map(f=>`<div class="frame">#${{f.index}} ${{esc(f.frame_type)}}/${{esc(f.frame_subtype)}}<br>${{esc(f.source)}} -> ${{esc(f.target)}}<br>start=${{f.start_sample}}, FCS=${{esc(f.fcs_ok)}}</div>`).join("")||"<p class='muted'>No frames.</p>"}}</div>`;
  renderList(); draw();
}}
function selectEdge(id){{
  selected=id; const e=edgeMap.get(id);
  const fs=graph.frames.filter(f=>e.frames.includes(f.index));
  details.innerHTML=`<div class="panel"><h2 style="font-size:16px;margin:0 0 10px">${{esc(e.source)}} -> ${{esc(e.target)}}</h2>`+
    rows([["Frames",e.weight],["SSID",e.ssid],["Relation",e.relation]])+
    `<h3>Frames</h3>${{fs.slice(0,25).map(f=>`<div class="frame">#${{f.index}} ${{esc(f.frame_type)}}/${{esc(f.frame_subtype)}} start=${{f.start_sample}} FCS=${{esc(f.fcs_ok)}}</div>`).join("")}}</div>`;
  renderList(); draw();
}}
function renderList(){{
  const q=filter.value.trim().toLowerCase();
  const list=graph.nodes.filter(n=>!q||[n.id,n.label,n.kind,n.ssid,n.bssid].join(" ").toLowerCase().includes(q));
  nodeList.innerHTML=list.map(n=>`<div class="item ${{selected===n.id?"active":""}}" data-id="${{esc(n.id)}}"><b>${{esc(n.label)}}</b><div class="muted">${{esc(n.kind)}} · ${{n.frames}} frames</div></div>`).join("")||`<div class="empty">No nodes.</div>`;
  nodeList.querySelectorAll(".item").forEach(el=>el.onclick=()=>selectNode(el.dataset.id));
}}
filter.oninput=renderList; window.onresize=draw; renderList(); draw();
</script>
</body>
</html>
"""

# 生成更完整的三栏交互式 index.html 页面，默认由 write_topology_html 写到脚本同目录。
def topology_html(graph: Dict[str, object]) -> str:
    graph_json = json.dumps(graph, ensure_ascii=False).replace("</", "<\\/")
    template = r"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>WiFi Signal Topology</title>
<style>
:root{--bg:#eef2f6;--panel:#fff;--soft:#f8fafc;--line:#d8e0ea;--text:#172033;--muted:#667085;--blue:#2563eb;--green:#079669;--purple:#7c3aed;--red:#dc2626}
*{box-sizing:border-box}
body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--text);letter-spacing:0}
.app{display:grid;grid-template-columns:316px minmax(520px,1fr) 392px;height:100vh;min-height:620px}
.side,.details{background:var(--panel);overflow:auto}
.side{border-right:1px solid var(--line)}
.details{border-left:1px solid var(--line)}
.brand{padding:20px 18px 14px;border-bottom:1px solid var(--line)}
.brandTop{display:flex;align-items:flex-start;justify-content:space-between;gap:10px}
h1{font-size:18px;line-height:1.2;margin:0 0 7px;font-weight:700}
.subtle{font-size:12px;color:var(--muted);line-height:1.45;word-break:break-all}
.settingsButton{width:32px;height:32px;display:inline-flex;align-items:center;justify-content:center;border:1px solid #cbd5e1;border-radius:7px;background:#fff;color:#475467;font-size:17px;line-height:1;cursor:pointer;flex:0 0 auto}
.settingsButton:hover{background:#f8fafc;border-color:#94a3b8;color:#172033}
.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;padding:14px 14px 12px}
.stat{background:var(--soft);border:1px solid var(--line);border-radius:8px;padding:10px 9px;min-width:0}
.stat b{display:block;font-size:19px;line-height:1;color:var(--text)}
.stat span{display:block;margin-top:5px;font-size:11px;color:var(--muted);white-space:nowrap}
.tools{padding:0 14px 12px;border-bottom:1px solid var(--line)}
.search{width:100%;height:36px;padding:7px 10px;border:1px solid #cbd5e1;border-radius:7px;background:#fff;font-size:13px;outline:none}
.search:focus{border-color:#7aa7ff;box-shadow:0 0 0 3px #dbeafe}
.legend{display:flex;flex-wrap:wrap;gap:7px;margin-top:10px}
.chip{display:inline-flex;align-items:center;gap:6px;min-height:24px;padding:3px 8px;border:1px solid var(--line);border-radius:999px;background:#fff;font-size:11px;color:#475467}
.dot{width:8px;height:8px;border-radius:50%;display:inline-block}
.nodeList{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:8px;padding:8px 8px 16px;align-items:start}
.nodeColumn{min-width:0}
.nodeColumnHead{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:8px 7px 7px;color:#475467;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em}
.nodeColumnHead b{font-size:11px;color:#667085;background:#f1f5f9;border-radius:999px;padding:2px 6px}
.nodeItem{width:100%;display:grid;grid-template-columns:12px minmax(0,1fr) auto;gap:7px;align-items:center;padding:8px 7px;border:1px solid #edf1f6;border-radius:8px;background:#fff;color:inherit;text-align:left;cursor:pointer;margin-bottom:7px;min-width:0}
.nodeItem:hover,.nodeItem.active{background:#eef6ff;border-color:#bfdbfe}
.nodeItem.sourceTarget{border-color:#d4a017;box-shadow:0 0 0 1px #d4a017;background:#fffdf5}
.nodeItem.sourceTarget:hover,.nodeItem.sourceTarget.active{background:#fff7d6;border-color:#d4a017}
.nodeItem.sourceOnly{border-color:#111827;box-shadow:0 0 0 1px #111827}
.nodeItem.sourceOnly:hover,.nodeItem.sourceOnly.active{background:#f8fafc;border-color:#111827}
.nodeItem.targetOnly{border-color:var(--node-color,#079669);border-style:dashed;box-shadow:0 0 0 1px transparent;background:#fff}
.nodeItem.targetOnly:hover,.nodeItem.targetOnly.active{background:#f8fffb;border-color:var(--node-color,#079669)}
.nodeItem .name{display:block;font-size:12px;font-weight:650;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.nodeItem .meta{display:block;font-size:10.5px;color:var(--muted);margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.nodeItem .count{font-size:11px;font-variant-numeric:tabular-nums;color:#475467;background:#f1f5f9;border-radius:6px;padding:2px 5px}
.swatch{width:11px;height:11px;border-radius:50%;box-shadow:0 0 0 3px #fff,0 0 0 4px #d8e0ea}
.stage{position:relative;overflow:hidden;background:linear-gradient(#ffffffcc,#ffffffcc),radial-gradient(circle at 20px 20px,#d8e0ea 1px,transparent 1px);background-size:auto,26px 26px}
.stageTop{position:absolute;left:18px;right:18px;top:14px;z-index:2;display:flex;justify-content:space-between;align-items:flex-start;gap:12px;pointer-events:none}
.stageTitle{background:#fffffff2;border:1px solid var(--line);border-radius:8px;padding:10px 12px;box-shadow:0 8px 24px #34405414}
.stageTitle b{display:block;font-size:13px;margin-bottom:4px}
.stageTitle span{font-size:12px;color:var(--muted)}
.badgeRow{display:flex;flex-wrap:wrap;justify-content:flex-end;gap:7px}
.badge{background:#fffffff2;border:1px solid var(--line);border-radius:999px;padding:6px 9px;font-size:11px;color:#475467;box-shadow:0 8px 24px #34405412}
svg{width:100%;height:100%;display:block;touch-action:none;cursor:grab}
svg.panning{cursor:grabbing}
.zoomControls{position:absolute;right:18px;bottom:18px;z-index:3;display:flex;gap:6px;background:#fffffff2;border:1px solid var(--line);border-radius:8px;padding:6px;box-shadow:0 8px 24px #34405418}
.zoomButton{width:32px;height:32px;border:1px solid #cbd5e1;border-radius:7px;background:#fff;color:#172033;font-size:14px;font-weight:700;line-height:1;cursor:pointer}
.zoomButton:hover{background:#f8fafc;border-color:#94a3b8}
.edge{stroke:#94a3b8;stroke-opacity:.66;cursor:pointer;fill:none}
.edge.related{stroke:#4f7cff;stroke-opacity:.9}
.edge.selected{stroke:var(--red);stroke-opacity:1}
.edge.faded{stroke-opacity:.16}
.edge.dashed{stroke-dasharray:8 7;stroke-linecap:round}
.edgeHit{stroke:transparent;stroke-width:16;cursor:pointer;fill:none}
.node{stroke:#fff;stroke-width:2.5;cursor:grab;filter:drop-shadow(0 6px 10px #0f172a22)}
.node.related{stroke:#172033;stroke-width:2.5}
.node.selected{stroke:#111827;stroke-width:4}
.node.faded{opacity:.28}
.node.targetOnly{fill:#fff;stroke:var(--node-color,#079669);stroke-width:2.8;stroke-dasharray:6 5;filter:none}
.node.targetOnly.related{stroke:var(--node-color,#079669);stroke-width:3}
.node.targetOnly.selected{stroke:#111827;stroke-width:4;stroke-dasharray:7 5}
.apNode{cursor:grab}
.apNode.faded{opacity:.28}
.iconNode{cursor:grab}
.iconNode.faded{opacity:.28}
.routerImage,.deviceImage{pointer-events:auto;cursor:grab;filter:drop-shadow(0 7px 10px #0f172a25)}
.apNode.selected .routerImage{filter:drop-shadow(0 8px 12px #11182755)}
.apNode.related .routerImage{filter:drop-shadow(0 8px 12px #17203335)}
.iconNode.selected .deviceImage{filter:drop-shadow(0 8px 12px #11182755)}
.iconNode.related .deviceImage{filter:drop-shadow(0 8px 12px #17203335)}
.targetOnlyRing{fill:none;stroke:var(--node-color,#2563eb);stroke-width:2.6;stroke-dasharray:7 5;pointer-events:none}
.componentCircle{fill:none;stroke:#94a3b8;stroke-width:1.7;stroke-dasharray:12 10;stroke-opacity:.62;pointer-events:none}
.label{font-size:11px;text-anchor:middle;fill:#1f2937;pointer-events:none;font-weight:650}
.kindLabel{font-size:9px;text-anchor:middle;fill:#667085;pointer-events:none}
.edgeLabel{font-size:11px;fill:#475467;text-anchor:middle;pointer-events:none;font-variant-numeric:tabular-nums}
.edgeLabelBg{fill:#fff;stroke:#d8e0ea;stroke-width:1}
.panel{padding:17px 16px 22px}
.panelHead{padding:18px 16px 14px;border-bottom:1px solid var(--line);background:#fff;position:sticky;top:0;z-index:3}
.panelHead h2{font-size:17px;line-height:1.25;margin:0 0 7px;word-break:break-all}
.typePill{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--line);border-radius:999px;padding:4px 9px;font-size:11px;color:#475467;background:#fff}
.metricGrid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin:14px 0}
.metric{background:var(--soft);border:1px solid var(--line);border-radius:8px;padding:9px}
.metric b{display:block;font-size:18px;line-height:1}
.metric span{display:block;margin-top:5px;font-size:11px;color:var(--muted)}
table{width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed}
td{border-bottom:1px solid #edf1f6;padding:8px 4px;vertical-align:top;word-break:break-word}
td:first-child{color:var(--muted);width:112px}
.sectionTitle{font-size:12px;text-transform:uppercase;color:#667085;letter-spacing:.04em;margin:18px 0 8px}
.frame{font-size:12px;border:1px solid #e7edf5;background:#fbfdff;border-radius:8px;padding:9px 10px;margin:7px 0;line-height:1.45}
.frame b{font-variant-numeric:tabular-nums}
.empty{padding:18px 16px;color:#667085;font-size:13px;line-height:1.5}
.actions{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0 4px}
.button{height:34px;border:1px solid #cbd5e1;border-radius:7px;background:#fff;color:#172033;padding:0 11px;font-size:12px;cursor:pointer}
.button:hover{background:#f8fafc;border-color:#94a3b8}
.modal{position:fixed;inset:0;background:#0f172a66;display:none;align-items:center;justify-content:center;padding:28px;z-index:10}
.modal.show{display:flex}
.modalShell{width:min(1180px,96vw);max-height:92vh;overflow-y:auto;overflow-x:hidden;background:#fff;border-radius:10px;box-shadow:0 24px 80px #0f172a66;border:1px solid #d8e0ea}
.modalTop{position:sticky;top:0;background:#fff;border-bottom:1px solid var(--line);padding:13px 15px;display:flex;justify-content:space-between;align-items:center;gap:12px;z-index:2}
.modalTop b{display:block;font-size:15px;word-break:break-all}
.modalTop span{font-size:12px;color:var(--muted)}
.plotGrid{display:grid;grid-template-columns:1fr 1fr;gap:14px;padding:14px}
.plotFigure{margin:0;border:1px solid #d8e0ea;border-radius:8px;background:#f8fafc;min-height:280px;overflow:hidden}
.plotFigure figcaption{font-size:12px;color:#475467;padding:10px 12px;border-bottom:1px solid #d8e0ea;background:#fff}
.plotFigure img{display:block;width:100%;height:auto;background:#fff}
.plotFigure .missingNote{display:none;margin:0;padding:28px 16px;color:#667085;font-size:13px}
.plotFigure.missing img{display:none}
.plotFigure.missing .missingNote{display:block}
.abbrTable{padding:14px}
.abbrTable table{font-size:13px}
.abbrTable th{font-size:12px;text-align:left;color:#475467;background:#f8fafc;border-bottom:1px solid #d8e0ea;padding:9px 8px}
.abbrTable td{padding:9px 8px}
.abbrTable td:first-child{width:88px;font-weight:700;color:#172033}
.abbrReferences{border-top:1px solid #d8e0ea;padding:14px 16px 18px;background:#fff}
.abbrReferences h3{font-size:13px;margin:0 0 10px;color:#172033}
.abbrReferences p{font-size:12px;line-height:1.55;margin:8px 0;color:#344054;word-break:break-word}
.abbrReferences .author{margin-top:12px;font-weight:700;color:#172033}
.channelWaveforms{border-top:1px solid #d8e0ea;padding:14px 16px 18px;background:#fbfdff;overflow:hidden}
.channelWaveforms h3{font-size:13px;margin:0 0 10px;color:#172033}
.channelWaveforms .channelNote{font-size:12px;color:#667085;margin:0 0 12px;line-height:1.45}
.channelGrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:10px;min-width:0}
.channelCard{margin:0;border:1px solid #d8e0ea;border-radius:8px;background:#fff;overflow:hidden;min-width:0;max-width:100%}
.channelCard figcaption{display:flex;justify-content:space-between;gap:8px;align-items:center;padding:7px 9px;border-bottom:1px solid #edf1f6;font-size:12px;color:#172033;min-width:0}
.channelCard figcaption b{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.channelCard figcaption span{font-size:11px;color:#667085;white-space:nowrap}
.channelCard img{display:block;width:100%;max-width:100%;height:auto;object-fit:contain;background:#fff}
.channelEmpty{font-size:12px;color:#667085;margin:0}
.frameLegend{border-top:1px solid #d8e0ea;padding:14px 16px 18px;background:#fff}
.frameLegend h3{font-size:13px;margin:0 0 10px;color:#172033}
.frameLegendGrid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}
.frameLegendItem{border:1px solid #e7edf5;border-radius:8px;background:#fbfdff;padding:10px;min-width:0}
.frameLegendSwatch{width:38px;height:26px;border-radius:7px;background:#fff;margin-bottom:8px}
.frameLegendSwatch.sourceOnly{border:2px solid #111827}
.frameLegendSwatch.sourceTarget{border:2px solid #d4a017}
.frameLegendSwatch.targetOnly{border:2px dashed #079669}
.frameLegendItem b{display:block;font-size:12px;color:#172033;margin-bottom:5px}
.frameLegendItem p{font-size:12px;line-height:1.45;color:#667085;margin:0}
@media (max-width:1500px){.app{grid-template-columns:300px minmax(0,1fr);grid-template-rows:720px 340px;height:auto;min-height:100vh}.side,.stage{height:720px}.details{grid-column:1/-1;height:340px;border-left:0;border-top:1px solid var(--line)}}
@media (max-width:760px){.app{display:block;height:auto}.side,.details{height:auto;max-height:none;border:0;border-bottom:1px solid var(--line)}.stage{height:620px}.plotGrid,.channelGrid,.frameLegendGrid{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="app">
<aside class="side">
  <div class="brand">
    <h1>WiFi Signal Topology</h1>
    <div class="subtle" id="sourcePath"></div>
  </div>
  <div class="stats" id="stats"></div>
  <div class="tools">
    <input class="search" id="filter" placeholder="Search MAC / SSID / type">
    <div class="legend" id="legend"></div>
  </div>
  <div class="nodeList" id="nodeList"></div>
</aside>
<main class="stage">
  <div class="stageTop">
    <div class="stageTitle"><b>Decoded Device Graph</b><span id="stageSummary"></span></div>
    <div class="badgeRow" id="badges"></div>
  </div>
  <div class="zoomControls" aria-label="Topology zoom controls">
    <button class="zoomButton" id="zoomOut" title="Zoom out">-</button>
    <button class="zoomButton" id="zoomReset" title="Reset view">1:1</button>
    <button class="zoomButton" id="zoomIn" title="Zoom in">+</button>
  </div>
  <svg id="graph"></svg>
</main>
<section class="details">
  <div class="panelHead"><h2>Topology Overview</h2><span class="typePill"><span class="dot" style="background:#94a3b8"></span>Overview</span></div>
  <div id="details" class="empty"></div>
</section>
</div>
<div class="modal" id="plotModal" aria-hidden="true">
  <div class="modalShell">
    <div class="modalTop">
      <div><b id="modalTitle"></b><span id="modalSub"></span></div>
      <button class="button" id="closeModal">Close</button>
    </div>
    <div class="plotGrid">
      <figure class="plotFigure"><figcaption>Extracted Signal Timeline</figcaption><img id="timelinePlot" alt=""><p class="missingNote">No timeline image found in signal_plots.</p></figure>
      <figure class="plotFigure"><figcaption>Single Frame Detail</figcaption><img id="framePlot" alt=""><p class="missingNote">No frame image found in signal_plots.</p></figure>
    </div>
  </div>
</div>
<script id="graph-data" type="application/json">__GRAPH_JSON__</script>
<script>
const graph = JSON.parse(document.getElementById("graph-data").textContent);
const svg = document.getElementById("graph");
const details = document.getElementById("details");
const stats = document.getElementById("stats");
const nodeList = document.getElementById("nodeList");
const filter = document.getElementById("filter");
const legend = document.getElementById("legend");
const badges = document.getElementById("badges");
const sourcePath = document.getElementById("sourcePath");
const stageSummary = document.getElementById("stageSummary");
const zoomInButton = document.getElementById("zoomIn");
const zoomOutButton = document.getElementById("zoomOut");
const zoomResetButton = document.getElementById("zoomReset");
let selectedType = "", selectedId = "";
let graphLayer = null;
let view = {scale:1,x:0,y:0};
let dragState = null;
let nodeDrag = null;
let suppressNodeClick = "";
const manualPos = new Map();
const routerIconPath = "assets/router_2p5d_clean_node.png";
const intelLaptopIconPath = "assets/laptop_2p5d_intel_node.png";
const nodeMap = new Map(graph.nodes.map(n => [n.id, n]));
const edgeMap = new Map(graph.edges.map(e => [e.id, e]));
const edgeOrder = new Map(graph.edges.map((e,i) => [e.id, i]));
const summary = graph.summary || {};
const adjacency = new Map();
for (const e of graph.edges) {
  if (!adjacency.has(e.source)) adjacency.set(e.source, new Set());
  if (!adjacency.has(e.target)) adjacency.set(e.target, new Set());
  adjacency.get(e.source).add(e.target);
  adjacency.get(e.target).add(e.source);
}
function esc(v){return String(v ?? "").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]));}
function colorKind(kind){return kind==="Access Point"?"#2563eb":kind==="Hidden AP"?"#7c3aed":kind==="Broadcast"?"#dc2626":kind==="Link"?"#94a3b8":"#079669";}
function shortMac(id){return id==="ff:ff:ff:ff:ff:ff"?kindText("Broadcast"):String(id||"").slice(-8);}
function cleanDisplayText(value){
  const raw=String(value??"");
  if(raw.trim()==="<hidden>")return "";
  let s=raw.replace(/<hidden>/gi," ").replace(/[\u0000-\u001f\u007f-\u009f\ufffd]/g," ");
  s=Array.from(s).filter(ch=>/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}A-Za-z0-9\s\-_.:@#[\]()+=~!$%^&,]/u.test(ch)).join("").replace(/\s+/g," ").trim();
  if(!s)return "";
  const chars=Array.from(s);
  const useful=chars.filter(ch=>/[\p{L}\p{N}]/u.test(ch)).length;
  if(!useful)return "";
  return s;
}
function ellipsize(value,max){const s=String(value??"");return s.length>max?s.slice(0,max-3)+"...":s;}
function displaySsid(value,fallback=""){const raw=String(value??"");const clean=cleanDisplayText(raw);return clean||((raw.trim()==="<hidden>")?"Hidden WiFi":fallback);}
function apTitle(n){return displaySsid(n.ssid,shortMac(n.id));}
function apLabel(n){return ellipsize(apTitle(n),26);}
function apIconSize(n){return Math.max(58,Math.min(78,54+Math.sqrt(Number(n.frames)||1)*3));}
function deviceIconSize(n){return Math.max(54,Math.min(76,50+Math.sqrt(Number(n.frames)||1)*2.8));}
function usesDeviceIcon(n){return n&&n.device_icon==="intel_laptop"&&!isApNode(n)&&n.kind!=="Broadcast";}
function nodeRadius(n){return isApNode(n)?apIconSize(n)*0.55:(usesDeviceIcon(n)?deviceIconSize(n)*0.48:Math.max(13,Math.min(32,10+Math.sqrt(Number(n.frames)||1)*2.6)));}
function degree(id){return adjacency.get(id)?.size||0;}
function kindRank(n){return n.kind==="Access Point"?0:n.kind==="Hidden AP"?1:n.kind==="Station"?2:3;}
function plotStem(id){return String(id||"node").toLowerCase().replace(/[^a-z0-9]+/g,"_").replace(/^_+|_+$/g,"")||"node";}
function plotPath(n,which){const key=which==="timeline"?"timeline_plot":"frame_plot";return n[key]||`signal_plots/${plotStem(n.id)}_${which}.png`;}
function makeSvg(tag){return document.createElementNS("http://www.w3.org/2000/svg",tag);}
function clamp(value,min,max){return Math.max(min,Math.min(max,value));}
function applyView(){if(graphLayer)graphLayer.setAttribute("transform",`translate(${view.x} ${view.y}) scale(${view.scale})`);}
function zoomAt(clientX,clientY,factor){
  const rect=svg.getBoundingClientRect();
  const sx=clientX-rect.left,sy=clientY-rect.top;
  const next=clamp(view.scale*factor,0.35,4);
  const graphX=(sx-view.x)/view.scale,graphY=(sy-view.y)/view.scale;
  view.x=sx-graphX*next;view.y=sy-graphY*next;view.scale=next;applyView();
}
function resetView(){view={scale:1,x:0,y:0};applyView();}
function graphPoint(evt){
  const rect=svg.getBoundingClientRect();
  return {x:(evt.clientX-rect.left-view.x)/view.scale,y:(evt.clientY-rect.top-view.y)/view.scale};
}
function applyManualPositions(pos){
  for(const [id,p] of manualPos.entries()){
    const target=pos.get(id);
    if(target){
      target.x=p.x;target.y=p.y;
      const guide=(pos.guides||[]).find(g=>Math.hypot(p.x-g.cx,p.y-g.cy)<=g.r+60)||pos.meta;
      target.angle=guide?angleNorm(Math.atan2((p.y-guide.cy)/(guide.outerRy||1),(p.x-guide.cx)/(guide.outerRx||1))):(p.angle??target.angle);
      target.ring=p.ring??target.ring;
    }
  }
}
function selectedNodeRelated(id){return selectedType!=="node" || id===selectedId || adjacency.get(selectedId)?.has(id);}
function edgeRelated(e){return !selectedType || (selectedType==="edge" && e.id===selectedId) || (selectedType==="node" && (e.source===selectedId || e.target===selectedId));}
function primaryAp(id){
  let best=null,bestWeight=-1;
  for(const e of graph.edges){
    if(e.source!==id && e.target!==id) continue;
    const other=nodeMap.get(e.source===id?e.target:e.source);
    if(!other || (other.kind!=="Access Point" && other.kind!=="Hidden AP")) continue;
    if(Number(e.weight)>bestWeight){best=other;bestWeight=Number(e.weight);}
  }
  return best;
}
function isApNode(n){return n && (n.kind==="Access Point"||n.kind==="Hidden AP");}
function isTargetOnlyNode(n){return !!n && (n.target_only===true || n.kind==="Broadcast" || ((Number(n.sent)||0)===0 && (Number(n.received)||0)>0));}
function isRadiatedNode(n){return !!n && n.has_radiated_signal===true && !isTargetOnlyNode(n);}
function edgeUsesDashedLine(e){return !(isRadiatedNode(nodeMap.get(e.source)) && isRadiatedNode(nodeMap.get(e.target)));}
function angleNorm(a){while(a<0)a+=Math.PI*2;while(a>=Math.PI*2)a-=Math.PI*2;return a;}
function angleDelta(from,to){
  let d=to-from;
  while(d>Math.PI)d-=Math.PI*2;
  while(d<-Math.PI)d+=Math.PI*2;
  return d;
}
function polarPoint(cx,cy,rx,ry,angle,ring){
  return {x:cx+Math.cos(angle)*rx,y:cy+Math.sin(angle)*ry,angle:angleNorm(angle),ring};
}
function placeOnEllipse(pos,n,cx,cy,rx,ry,angle,ring){
  pos.set(n.id,polarPoint(cx,cy,rx,ry,angle,ring));
}
function nodeSort(a,b){return Number(b.frames||0)-Number(a.frames||0)||String(a.id).localeCompare(String(b.id));}
function apNeighbors(id){
  const neighbors=[];
  for(const e of graph.edges){
    if(e.source!==id && e.target!==id)continue;
    const other=nodeMap.get(e.source===id?e.target:e.source);
    if(isApNode(other))neighbors.push(other);
  }
  return neighbors;
}
function connectedApCenterAngle(node,apAngle){
  const aps=apNeighbors(node.id);
  if(!aps.length)return null;
  let sx=0,sy=0;
  for(const ap of aps){
    const a=apAngle.get(ap.id);
    if(typeof a==="number"){sx+=Math.cos(a);sy+=Math.sin(a);}
  }
  return Math.atan2(sy,sx);
}
function relaxCollisions(pos,w,h,margin){
  for(let iter=0;iter<90;iter++){
    let moved=false;
    for(let i=0;i<graph.nodes.length;i++)for(let j=i+1;j<graph.nodes.length;j++){
      const a=graph.nodes[i],b=graph.nodes[j],pa=pos.get(a.id),pb=pos.get(b.id);if(!pa||!pb)continue;
      let dx=pb.x-pa.x,dy=pb.y-pa.y,dist=Math.sqrt(dx*dx+dy*dy)||0.01;
      const minDist=nodeRadius(a)+nodeRadius(b)+34;
      if(dist<minDist){
        const push=(minDist-dist)/2+0.2;dx/=dist;dy/=dist;
        pa.x-=dx*push;pa.y-=dy*push;pb.x+=dx*push;pb.y+=dy*push;moved=true;
      }
    }
    for(const n of graph.nodes){
      const p=pos.get(n.id);if(!p)continue;
      p.x=Math.max(margin,Math.min(w-margin,p.x));
      p.y=Math.max(margin+28,Math.min(h-margin,p.y));
    }
    if(!moved)break;
  }
}
function refreshPolarAngles(pos){
  const guides=pos.guides||[];if(!guides.length&&!pos.meta)return;
  for(const n of graph.nodes){
    const p=pos.get(n.id);if(!p)continue;
    const guide=guides.find(g=>Math.hypot(p.x-g.cx,p.y-g.cy)<=g.r+80)||pos.meta;
    if(guide)p.angle=angleNorm(Math.atan2((p.y-guide.cy)/(guide.outerRy||1),(p.x-guide.cx)/(guide.outerRx||1)));
  }
}
function pointClose(a,b){return Math.hypot(a.x-b.x,a.y-b.y)<1.5;}
function orient(a,b,c){
  const value=(b.x-a.x)*(c.y-a.y)-(b.y-a.y)*(c.x-a.x);
  return Math.abs(value)<0.001?0:(value>0?1:-1);
}
function onSegment(a,b,c){
  return Math.min(a.x,c.x)-0.001<=b.x&&b.x<=Math.max(a.x,c.x)+0.001&&
         Math.min(a.y,c.y)-0.001<=b.y&&b.y<=Math.max(a.y,c.y)+0.001;
}
function segmentsCross(a,b,c,d){
  if(pointClose(a,c)||pointClose(a,d)||pointClose(b,c)||pointClose(b,d))return false;
  const o1=orient(a,b,c),o2=orient(a,b,d),o3=orient(c,d,a),o4=orient(c,d,b);
  if(o1!==o2&&o3!==o4)return true;
  return (o1===0&&onSegment(a,c,b))||(o2===0&&onSegment(a,d,b))||
         (o3===0&&onSegment(c,a,d))||(o4===0&&onSegment(c,b,d));
}
function edgeSegments(pos,e){
  const a=pos.get(e.source),b=pos.get(e.target);if(!a||!b)return [];
  const pts=routedEdgePoints(e,a,b,pos),segments=[];
  for(let i=0;i<pts.length-1;i++)segments.push([pts[i],pts[i+1]]);
  return segments;
}
function routesCross(pos,a,b){
  const aSegs=edgeSegments(pos,a),bSegs=edgeSegments(pos,b);
  for(const sa of aSegs)for(const sb of bSegs)if(segmentsCross(sa[0],sa[1],sb[0],sb[1]))return true;
  return false;
}
function layoutScore(pos){
  let crossings=0,overlap=0,lengthSum=0;
  for(let i=0;i<graph.edges.length;i++)for(let j=i+1;j<graph.edges.length;j++){
    const a=graph.edges[i],b=graph.edges[j];
    if(a.source===b.source||a.source===b.target||a.target===b.source||a.target===b.target)continue;
    if(routesCross(pos,a,b))crossings++;
  }
  for(let i=0;i<graph.nodes.length;i++)for(let j=i+1;j<graph.nodes.length;j++){
    const a=graph.nodes[i],b=graph.nodes[j],pa=pos.get(a.id),pb=pos.get(b.id);if(!pa||!pb)continue;
    const dist=Math.hypot(pb.x-pa.x,pb.y-pa.y),minDist=nodeRadius(a)+nodeRadius(b)+30;
    if(dist<minDist)overlap+=minDist-dist;
  }
  for(const e of graph.edges){
    const a=pos.get(e.source),b=pos.get(e.target);if(a&&b)lengthSum+=Math.hypot(b.x-a.x,b.y-a.y);
  }
  return crossings*100000+overlap*1000+lengthSum/1000;
}
function untangleLayout(pos){
  const swappable=graph.nodes.filter(n=>!isApNode(n)&&n.kind!=="Broadcast"&&pos.has(n.id));
  let best=layoutScore(pos);
  for(let pass=0;pass<5;pass++){
    let changed=false;
    for(let i=0;i<swappable.length;i++)for(let j=i+1;j<swappable.length;j++){
      const a=pos.get(swappable[i].id),b=pos.get(swappable[j].id);if(!a||!b)continue;
      if(a.ring!=="outer"||b.ring!=="outer")continue;
      const ax=a.x,ay=a.y,aa=a.angle,bx=b.x,by=b.y,ba=b.angle;
      a.x=bx;a.y=by;a.angle=ba;b.x=ax;b.y=ay;b.angle=aa;
      const score=layoutScore(pos);
      if(score+0.001<best){best=score;changed=true;}else{a.x=ax;a.y=ay;a.angle=aa;b.x=bx;b.y=by;b.angle=ba;}
    }
    if(!changed)break;
  }
  return pos;
}
function makeRingLayout(reverseAp,rotate,freeRotate){
  const w=svg.clientWidth||920,h=svg.clientHeight||680,margin=82;
  const cx=w/2,cy=h/2+20;
  const outerRx=Math.max(170,w/2-margin-18),outerRy=Math.max(210,h/2-margin-42);
  const innerRx=Math.max(88,outerRx*0.48),innerRy=Math.max(78,outerRy*0.42);
  const aps=graph.nodes.filter(isApNode).sort(nodeSort);
  if(reverseAp)aps.reverse();
  const broadcasts=graph.nodes.filter(n=>n.kind==="Broadcast").sort(nodeSort);
  const stations=graph.nodes.filter(n=>!isApNode(n)&&n.kind!=="Broadcast").sort(nodeSort);
  const pos=new Map();
  pos.meta={cx,cy,innerRx,innerRy,outerRx,outerRy,outerPad:46};
  const apAngle=new Map();
  const apCount=Math.max(aps.length,1);
  const base=-Math.PI/2+rotate;
  aps.forEach((ap,i)=>{
    const angle=base+i*Math.PI*2/apCount;
    apAngle.set(ap.id,angleNorm(angle));
    placeOnEllipse(pos,ap,cx,cy,innerRx,innerRy,angle,"inner");
  });
  const groups=new Map(aps.map(ap=>[ap.id,[]]));
  const free=[];
  for(const n of stations){
    const ap=primaryAp(n.id);
    if(ap&&groups.has(ap.id))groups.get(ap.id).push(n);
    else free.push(n);
  }
  const sector=Math.min(Math.PI*1.05,Math.PI*2/Math.max(apCount,1)*0.72);
  for(const ap of aps){
    const list=(groups.get(ap.id)||[]).sort((a,b)=>{
      const aa=connectedApCenterAngle(a,apAngle),bb=connectedApCenterAngle(b,apAngle);
      const da=aa===null?0:angleDelta(apAngle.get(ap.id),aa);
      const db=bb===null?0:angleDelta(apAngle.get(ap.id),bb);
      return da-db||nodeSort(a,b);
    });
    list.forEach((n,i)=>{
      const count=list.length;
      const angle=apAngle.get(ap.id)+(count===1?0:-sector/2+sector*(i+0.5)/count);
      placeOnEllipse(pos,n,cx,cy,outerRx,outerRy,angle,"outer");
    });
  }
  const freeCenter=angleNorm(Math.PI+freeRotate);
  const freeWidth=Math.min(Math.PI*0.9,Math.PI*2/Math.max(apCount,4)*0.95);
  const freeList=free.sort(nodeSort);
  freeList.forEach((n,i)=>{
    const count=freeList.length;
    const angle=freeCenter+(count===1?0:-freeWidth/2+freeWidth*(i+0.5)/count);
    placeOnEllipse(pos,n,cx,cy,outerRx,outerRy,angle,"outer");
  });
  broadcasts.forEach((n,i)=>{
    const count=broadcasts.length;
    const angle=angleNorm(freeCenter+Math.PI+(count===1?0:(i-(count-1)/2)*0.22));
    placeOnEllipse(pos,n,cx,cy,outerRx,outerRy,angle,"outer");
  });
  relaxCollisions(pos,w,h,margin);
  refreshPolarAngles(pos);
  return untangleLayout(pos);
}
function nodeDegreeCount(n){return adjacency.get(n.id)?.size||0;}
function connectedComponents(){
  const seen=new Set(),components=[];
  for(const n of [...graph.nodes].sort(nodeSort)){
    if(seen.has(n.id))continue;
    const stack=[n.id],ids=[];seen.add(n.id);
    while(stack.length){
      const id=stack.pop();ids.push(id);
      for(const next of adjacency.get(id)||[]){
        if(!seen.has(next)){seen.add(next);stack.push(next);}
      }
    }
    components.push(ids.map(id=>nodeMap.get(id)).filter(Boolean).sort((a,b)=>kindRank(a)-kindRank(b)||nodeSort(a,b)));
  }
  return components.sort((a,b)=>b.length-a.length||String(a[0]?.id||"").localeCompare(String(b[0]?.id||"")));
}
function componentEdges(nodes){
  const ids=new Set(nodes.map(n=>n.id));
  return graph.edges.filter(e=>ids.has(e.source)&&ids.has(e.target));
}
function componentRadius(nodes){
  const count=nodes.length,apCount=nodes.filter(isApNode).length,edgeCount=componentEdges(nodes).length;
  if(count<=1)return isApNode(nodes[0])?78:68;
  if(count===2)return 118;
  return Math.max(150,Math.min(330,72+Math.sqrt(count)*54+Math.sqrt(edgeCount)*12+Math.min(apCount,6)*8));
}
function packComponentCircles(components,w,h){
  const side=54,top=h>560?104:84,bottom=48,gap=34;
  const usableW=Math.max(320,w-side*2),usableH=Math.max(280,h-top-bottom);
  const items=components.map((nodes,index)=>({nodes,index,r:componentRadius(nodes)}));
  const maxR=Math.max(...items.map(x=>x.r),1);
  const totalArea=items.reduce((sum,x)=>sum+Math.PI*x.r*x.r,0);
  const areaScale=Math.min(1,Math.sqrt((usableW*usableH*0.58)/Math.max(totalArea,1)));
  const widthScale=Math.min(1,usableW/(maxR*2+gap));
  const scale=Math.max(0.52,Math.min(areaScale,widthScale));
  for(const item of items)item.r=Math.max(54,item.r*scale);
  const ordered=[...items].sort((a,b)=>b.r-a.r||a.index-b.index);
  const placed=[];
  let row=[],rowW=0,rowH=0,rows=[];
  for(const item of ordered){
    const d=item.r*2;
    if(row.length && rowW+d+gap>usableW){
      rows.push({items:row,width:rowW-gap,height:rowH});
      row=[];rowW=0;rowH=0;
    }
    row.push(item);rowW+=d+gap;rowH=Math.max(rowH,d);
  }
  if(row.length)rows.push({items:row,width:rowW-gap,height:rowH});
  const totalH=rows.reduce((sum,r)=>sum+r.height,0)+gap*Math.max(0,rows.length-1);
  let y=top+Math.max(0,(usableH-totalH)/2);
  for(const rowInfo of rows){
    let x=side+Math.max(0,(usableW-rowInfo.width)/2);
    for(const item of rowInfo.items){
      placed[item.index]={cx:x+item.r,cy:y+rowInfo.height/2,r:item.r};
      x+=item.r*2+gap;
    }
    y+=rowInfo.height+gap;
  }
  return placed;
}
function componentGuide(circle,nodeCount){
  const r=circle.r;
  const innerScale=nodeCount>6?0.48:0.42;
  const outerR=Math.max(44,r-48);
  return {
    cx:circle.cx,cy:circle.cy,r,
    innerRx:Math.max(42,outerR*innerScale),innerRy:Math.max(42,outerR*innerScale),
    outerRx:outerR,outerRy:outerR,outerPad:46,
    bounds:{x:circle.cx-r,y:circle.cy-r,width:r*2,height:r*2}
  };
}
function localAdjacency(nodes,edges){
  const local=new Map(nodes.map(n=>[n.id,new Set()]));
  for(const e of edges){local.get(e.source)?.add(e.target);local.get(e.target)?.add(e.source);}
  return local;
}
function pickComponentHub(nodes,local,reverseHub){
  const ranked=[...nodes].sort((a,b)=>(local.get(b.id)?.size||0)-(local.get(a.id)?.size||0)||kindRank(a)-kindRank(b)||nodeSort(a,b));
  if(reverseHub)ranked.reverse();
  return ranked[0]||nodes[0];
}
function clampToCircle(p,guide,pad=34){
  const limit=Math.max(20,guide.r-pad);
  const dx=p.x-guide.cx,dy=p.y-guide.cy,dist=Math.hypot(dx,dy);
  if(dist>limit){
    const scale=limit/Math.max(dist,0.001);
    p.x=guide.cx+dx*scale;
    p.y=guide.cy+dy*scale;
  }
}
function relaxComponentCollisions(pos,nodes,guide,lockedIds){
  for(let iter=0;iter<80;iter++){
    let moved=false;
    for(let i=0;i<nodes.length;i++)for(let j=i+1;j<nodes.length;j++){
      const a=nodes[i],b=nodes[j],pa=pos.get(a.id),pb=pos.get(b.id);if(!pa||!pb)continue;
      let dx=pb.x-pa.x,dy=pb.y-pa.y,dist=Math.sqrt(dx*dx+dy*dy)||0.01;
      const minDist=nodeRadius(a)+nodeRadius(b)+34;
      if(dist>=minDist)continue;
      dx/=dist;dy/=dist;const push=(minDist-dist)+0.2;
      const lockA=lockedIds.has(a.id),lockB=lockedIds.has(b.id);
      if(lockA&&lockB)continue;
      if(lockA){pb.x+=dx*push;pb.y+=dy*push;}
      else if(lockB){pa.x-=dx*push;pa.y-=dy*push;}
      else{pa.x-=dx*push/2;pa.y-=dy*push/2;pb.x+=dx*push/2;pb.y+=dy*push/2;}
      moved=true;
    }
    for(const n of nodes){const p=pos.get(n.id);if(p)clampToCircle(p,guide);}
    if(!moved)break;
  }
}
function placeComponent(pos,nodes,circle,rotate,reverseHub){
  const edges=componentEdges(nodes),local=localAdjacency(nodes,edges),guide=componentGuide(circle,nodes.length);
  pos.guides.push(guide);
  if(!nodes.length)return;
  const hub=pickComponentHub(nodes,local,reverseHub);
  pos.set(hub.id,{x:guide.cx,y:guide.cy,angle:0,ring:"center"});
  const hubNeighbors=[...(local.get(hub.id)||[])].map(id=>nodeMap.get(id)).filter(Boolean).sort((a,b)=>(local.get(b.id)?.size||0)-(local.get(a.id)?.size||0)||kindRank(a)-kindRank(b)||nodeSort(a,b));
  const placed=new Set([hub.id]);
  const hasSecondRing=nodes.some(n=>n.id!==hub.id&&!hubNeighbors.some(m=>m.id===n.id));
  const nearRx=hasSecondRing?guide.innerRx:guide.outerRx,nearRy=hasSecondRing?guide.innerRy:guide.outerRy;
  hubNeighbors.forEach((n,i)=>{
    const count=hubNeighbors.length;
    const angle=rotate-Math.PI/2+Math.PI*2*i/Math.max(count,1);
    placeOnEllipse(pos,n,guide.cx,guide.cy,nearRx,nearRy,angle,hasSecondRing?"inner":"outer");
    placed.add(n.id);
  });
  const anchorAngle=new Map(hubNeighbors.map(n=>[n.id,pos.get(n.id).angle]));
  anchorAngle.set(hub.id,rotate-Math.PI/2);
  const groups=new Map();
  for(const n of nodes){
    if(placed.has(n.id))continue;
    const neighbors=local.get(n.id)||new Set();
    let anchor=hubNeighbors.find(x=>neighbors.has(x.id))||hub;
    if(!groups.has(anchor.id))groups.set(anchor.id,[]);
    groups.get(anchor.id).push(n);
  }
  for(const [anchorId,list] of groups.entries()){
    list.sort((a,b)=>(local.get(b.id)?.size||0)-(local.get(a.id)?.size||0)||kindRank(a)-kindRank(b)||nodeSort(a,b));
    const base=anchorAngle.get(anchorId)??rotate;
    const width=list.length>3?Math.PI*0.95:Math.PI*0.62;
    list.forEach((n,i)=>{
      const angle=base+(list.length===1?0:-width/2+width*(i+0.5)/list.length);
      placeOnEllipse(pos,n,guide.cx,guide.cy,guide.outerRx,guide.outerRy,angle,"outer");
      placed.add(n.id);
    });
  }
  const leftovers=nodes.filter(n=>!placed.has(n.id)).sort(nodeSort);
  leftovers.forEach((n,i)=>{
    const angle=rotate+Math.PI*2*i/Math.max(leftovers.length,1);
    placeOnEllipse(pos,n,guide.cx,guide.cy,guide.outerRx,guide.outerRy,angle,"outer");
  });
  const locked=new Set([hub.id]);
  relaxComponentCollisions(pos,nodes,guide,locked);
  pos.set(hub.id,{x:guide.cx,y:guide.cy,angle:0,ring:"center"});
}
function makeComponentLayout(reverseHub,rotate){
  const w=svg.clientWidth||920,h=svg.clientHeight||680;
  const components=connectedComponents(),circles=packComponentCircles(components,w,h);
  const pos=new Map();pos.guides=[];pos.meta=null;
  components.forEach((nodes,i)=>placeComponent(pos,nodes,circles[i],rotate+i*Math.PI/7,reverseHub&&i%2===1));
  pos.meta=pos.guides[0]||null;
  refreshPolarAngles(pos);
  return pos;
}
function layout(){
  const candidates=[
    makeComponentLayout(false,0),makeComponentLayout(false,Math.PI/4),
    makeComponentLayout(false,Math.PI/2),makeComponentLayout(true,0),
    makeComponentLayout(true,Math.PI/4),makeComponentLayout(true,Math.PI/2)
  ];
  const pos=candidates.sort((a,b)=>layoutScore(a)-layoutScore(b))[0];
  applyManualPositions(pos);
  return pos;
}
function pathFromPoints(points){
  return `M ${points[0].x} ${points[0].y} `+points.slice(1).map(p=>`L ${p.x} ${p.y}`).join(" ");
}
function edgeArcPath(a,b,pos,outerGap=42){
  const meta=pos?.meta;
  if(!meta)return pathFromPoints([a,b]);
  const rxa=meta.outerRx+outerGap,rya=meta.outerRy+outerGap;
  const aa=a.angle??angleNorm(Math.atan2((a.y-meta.cy)/meta.outerRy,(a.x-meta.cx)/meta.outerRx));
  const bb=b.angle??angleNorm(Math.atan2((b.y-meta.cy)/meta.outerRy,(b.x-meta.cx)/meta.outerRx));
  const delta=Math.abs(angleDelta(aa,bb));
  const large=delta>Math.PI?1:0;
  const sweep=angleDelta(aa,bb)>0?1:0;
  const start=polarPoint(meta.cx,meta.cy,rxa,rya,aa,"arc");
  const end=polarPoint(meta.cx,meta.cy,rxa,rya,bb,"arc");
  return `M ${a.x} ${a.y} L ${start.x} ${start.y} A ${rxa} ${rya} 0 ${large} ${sweep} ${end.x} ${end.y} L ${b.x} ${b.y}`;
}
function edgeTouchesBroadcast(e){
  return nodeMap.get(e.source)?.kind==="Broadcast"||nodeMap.get(e.target)?.kind==="Broadcast";
}
function shouldRouteOutside(e,a,b,pos){
  if(!pos||edgeTouchesBroadcast(e))return false;
  const idx=edgeOrder.get(e.id)??graph.edges.indexOf(e);
  for(let i=0;i<idx;i++){
    const other=graph.edges[i];
    if(edgeTouchesBroadcast(other))continue;
    if(e.source===other.source||e.source===other.target||e.target===other.source||e.target===other.target)continue;
    const c=pos.get(other.source),d=pos.get(other.target);if(!c||!d)continue;
    if(segmentsCross(a,b,c,d))return true;
  }
  return false;
}
function outsideRoute(e,a,b,pos){
  const w=svg.clientWidth||920;
  const side=(a.x+b.x)/2 < w*0.48 ? -1 : 1;
  const edgeIndex=edgeOrder.get(e.id)??0;
  const offset=54+(edgeIndex%5)*18;
  const rawLane=side<0?Math.min(a.x,b.x)-offset:Math.max(a.x,b.x)+offset;
  const lane=side<0?Math.max(30,rawLane):Math.min(w-30,rawLane);
  return [a,{x:lane,y:a.y},{x:lane,y:b.y},b];
}
function routedEdgePoints(e,a,b,pos){
  const source=nodeMap.get(e.source),target=nodeMap.get(e.target);
  const sourceIsAp=source?.kind==="Access Point"||source?.kind==="Hidden AP";
  const targetIsAp=target?.kind==="Access Point"||target?.kind==="Hidden AP";
  if(edgeTouchesBroadcast(e))return [a,b];
  if(sourceIsAp!==targetIsAp)return [a,b];
  return [a,b];
}
function edgePathData(e,a,b,pos){
  return pathFromPoints([a,b]);
}
function edgeMidpoint(e,a,b,pos){
  return {x:(a.x+b.x)/2,y:(a.y+b.y)/2-6};
}
function startNodeDrag(evt,id,p){
  if(evt.button!==0)return;
  evt.stopPropagation();
  nodeDrag={id,pointerId:evt.pointerId,startClientX:evt.clientX,startClientY:evt.clientY,startX:p.x,startY:p.y,moved:false,lastClick:0};
  svg.setPointerCapture(evt.pointerId);svg.classList.add("panning");
}
function finishNodeDrag(evt,id){
  if(suppressNodeClick===id){suppressNodeClick="";evt.preventDefault();evt.stopPropagation();return false;}
  if(nodeDrag&&nodeDrag.id===id&&nodeDrag.moved){evt.preventDefault();evt.stopPropagation();return false;}
  return true;
}
function draw(){
  const pos=layout();
  svg.innerHTML="";
  graphLayer=makeSvg("g");svg.appendChild(graphLayer);applyView();
  for(const guide of pos.guides||[]){
    const c=makeSvg("circle");
    c.setAttribute("cx",guide.cx);c.setAttribute("cy",guide.cy);c.setAttribute("r",guide.r);
    c.setAttribute("class","componentCircle");
    graphLayer.appendChild(c);
  }
  for(const e of graph.edges){
    const a=pos.get(e.source),b=pos.get(e.target);if(!a||!b)continue;
    const d=edgePathData(e,a,b,pos);
    const hit=makeSvg("path");hit.setAttribute("d",d);hit.setAttribute("class","edgeHit");hit.dataset.edgeId=e.id;hit.dataset.source=e.source;hit.dataset.target=e.target;hit.addEventListener("click",()=>selectEdge(e.id));graphLayer.appendChild(hit);
    const line=makeSvg("path");line.setAttribute("d",d);line.setAttribute("stroke-width",Math.max(1.5,Math.min(8,1.2+Math.sqrt(Number(e.weight)||1))));
    const classes=["edge"];if(edgeUsesDashedLine(e))classes.push("dashed");if(selectedType==="edge"&&selectedId===e.id)classes.push("selected");else if(edgeRelated(e))classes.push("related");else classes.push("faded");line.setAttribute("class",classes.join(" "));line.dataset.edgeId=e.id;line.dataset.source=e.source;line.dataset.target=e.target;line.addEventListener("click",()=>selectEdge(e.id));graphLayer.appendChild(line);
    if((Number(e.weight)||0)>1 || (selectedType==="edge"&&selectedId===e.id)){const mid=edgeMidpoint(e,a,b,pos),mx=mid.x,my=mid.y,bg=makeSvg("rect"),t=makeSvg("text");bg.setAttribute("x",mx-12);bg.setAttribute("y",my-12);bg.setAttribute("width",24);bg.setAttribute("height",16);bg.setAttribute("rx",5);bg.setAttribute("class","edgeLabelBg");t.setAttribute("x",mx);t.setAttribute("y",my);t.setAttribute("class","edgeLabel");t.textContent=e.weight;graphLayer.appendChild(bg);graphLayer.appendChild(t);}
  }
  for(const n of graph.nodes){
    const p=pos.get(n.id);if(!p)continue;const r=nodeRadius(n);
    if(isApNode(n)){
      const size=apIconSize(n),g=makeSvg("g"),img=makeSvg("image");
      const classes=["apNode"];if(selectedType==="node"&&selectedId===n.id)classes.push("selected");else if(selectedNodeRelated(n.id))classes.push("related");else classes.push("faded");
      if(isTargetOnlyNode(n))classes.push("targetOnly");
      g.setAttribute("class",classes.join(" "));g.style.setProperty("--node-color",colorKind(n.kind));g.dataset.nodeId=n.id;g.addEventListener("pointerdown",evt=>startNodeDrag(evt,n.id,p));g.addEventListener("click",evt=>{if(finishNodeDrag(evt,n.id))selectNode(n.id);});g.addEventListener("dblclick",evt=>{if(finishNodeDrag(evt,n.id))openPlots(n.id);});graphLayer.appendChild(g);
      if(isTargetOnlyNode(n)){const ring=makeSvg("circle");ring.setAttribute("cx",p.x);ring.setAttribute("cy",p.y);ring.setAttribute("r",size*0.58);ring.setAttribute("class","targetOnlyRing");g.appendChild(ring);}
      img.setAttribute("href",routerIconPath);img.setAttribute("x",p.x-size*0.66);img.setAttribute("y",p.y-size*0.48);img.setAttribute("width",size*1.32);img.setAttribute("height",size*0.9);img.setAttribute("class","routerImage");g.appendChild(img);
      const title=makeSvg("title");title.textContent=`${apTitle(n)} | ${n.kind} | ${n.frames||0} frames`;g.appendChild(title);
      const label=makeSvg("text");label.setAttribute("x",p.x);label.setAttribute("y",p.y+size*0.55);label.setAttribute("class","label");label.textContent=apLabel(n);graphLayer.appendChild(label);
      const kind=makeSvg("text");kind.setAttribute("x",p.x);kind.setAttribute("y",p.y+size*0.72);kind.setAttribute("class","kindLabel");kind.textContent=n.kind==="Hidden AP"?"Hidden AP":"AP";graphLayer.appendChild(kind);
      continue;
    }
    if(usesDeviceIcon(n)){
      const size=deviceIconSize(n),g=makeSvg("g"),img=makeSvg("image"),nodeColor=colorKind(n.kind),targetOnly=isTargetOnlyNode(n);
      const classes=["iconNode"];if(selectedType==="node"&&selectedId===n.id)classes.push("selected");else if(selectedNodeRelated(n.id))classes.push("related");else classes.push("faded");
      if(targetOnly)classes.push("targetOnly");
      g.setAttribute("class",classes.join(" "));g.style.setProperty("--node-color",nodeColor);g.dataset.nodeId=n.id;g.addEventListener("pointerdown",evt=>startNodeDrag(evt,n.id,p));g.addEventListener("click",evt=>{if(finishNodeDrag(evt,n.id))selectNode(n.id);});g.addEventListener("dblclick",evt=>{if(finishNodeDrag(evt,n.id))openPlots(n.id);});graphLayer.appendChild(g);
      if(targetOnly){const ring=makeSvg("circle");ring.setAttribute("cx",p.x);ring.setAttribute("cy",p.y);ring.setAttribute("r",size*0.58);ring.setAttribute("class","targetOnlyRing");g.appendChild(ring);}
      img.setAttribute("href",intelLaptopIconPath);img.setAttribute("x",p.x-size*0.68);img.setAttribute("y",p.y-size*0.5);img.setAttribute("width",size*1.36);img.setAttribute("height",size);img.setAttribute("class","deviceImage");g.appendChild(img);
      const title=makeSvg("title");title.textContent=`${n.label||n.id} | Intel Station | ${n.frames||0} frames`;g.appendChild(title);
      const label=makeSvg("text");label.setAttribute("x",p.x);label.setAttribute("y",p.y+size*0.58);label.setAttribute("class","label");label.textContent=shortMac(n.id);graphLayer.appendChild(label);
      const kind=makeSvg("text");kind.setAttribute("x",p.x);kind.setAttribute("y",p.y+size*0.72);kind.setAttribute("class","kindLabel");kind.textContent="Intel Station";graphLayer.appendChild(kind);
      continue;
    }
    const c=makeSvg("circle");
    const nodeColor=colorKind(n.kind),targetOnly=isTargetOnlyNode(n);
    c.setAttribute("cx",p.x);c.setAttribute("cy",p.y);c.setAttribute("r",r);c.setAttribute("fill",targetOnly?"#fff":nodeColor);
    const classes=["node"];if(selectedType==="node"&&selectedId===n.id)classes.push("selected");else if(selectedNodeRelated(n.id))classes.push("related");else classes.push("faded");
    if(targetOnly)classes.push("targetOnly");
    c.setAttribute("class",classes.join(" "));c.style.setProperty("--node-color",nodeColor);c.dataset.nodeId=n.id;c.addEventListener("pointerdown",evt=>startNodeDrag(evt,n.id,p));c.addEventListener("click",evt=>{if(finishNodeDrag(evt,n.id))selectNode(n.id);});c.addEventListener("dblclick",evt=>{if(finishNodeDrag(evt,n.id))openPlots(n.id);});graphLayer.appendChild(c);
    const title=makeSvg("title");title.textContent=`${n.label||n.id} | ${n.kind} | ${n.frames||0} frames`;c.appendChild(title);
    const label=makeSvg("text");label.setAttribute("x",p.x);label.setAttribute("y",p.y+r+15);label.setAttribute("class","label");label.textContent=shortMac(n.id);graphLayer.appendChild(label);
    const kind=makeSvg("text");kind.setAttribute("x",p.x);kind.setAttribute("y",p.y+r+28);kind.setAttribute("class","kindLabel");kind.textContent=n.kind==="Broadcast"?"Broadcast":"Station";graphLayer.appendChild(kind);
  }
}
function rows(items){return `<table>${items.map(([k,v])=>`<tr><td>${esc(k)}</td><td>${esc(v)}</td></tr>`).join("")}</table>`;}
function metrics(items){return `<div class="metricGrid">${items.map(([k,v])=>`<div class="metric"><b>${esc(v)}</b><span>${esc(k)}</span></div>`).join("")}</div>`;}
function nodeFrames(id){return graph.frames.filter(f=>[f.source,f.target,String(f.bssid||"").toLowerCase(),String(f.receiver||"").toLowerCase(),String(f.transmitter||"").toLowerCase()].includes(id)).slice(0,22);}
function frameList(frames){return frames.map(f=>`<div class="frame"><b>#${esc(f.index)}</b> ${esc(f.frame_type)}/${esc(f.frame_subtype)}<br>${esc(f.source||"unknown")} -> ${esc(f.target||"unknown")}<br>start=${esc(f.start_sample)}, FCS=${esc(f.fcs_ok)}, MCS=${esc(f.mcs)}</div>`).join("")||`<div class="empty">${esc(tr("No frames."))}</div>`;}
function setDetailHeader(title,kind){document.querySelector(".panelHead").innerHTML=`<h2>${esc(title)}</h2><span class="typePill"><span class="dot" style="background:${colorKind(kind)}"></span>${esc(kind||"Selection")}</span>`;}
function selectNode(id){
  selectedType="node";selectedId=id;const n=nodeMap.get(id);if(!n)return;setDetailHeader(n.label||n.id,n.kind);details.className="panel";
  details.innerHTML=metrics([["Frames",n.frames||0],["Sent",n.sent||0],["Received",n.received||0]])+
    `<div class="actions"><button class="button" id="openPlotsButton">Open plots</button></div><div class="sectionTitle">Device</div>`+
    rows([["Type",n.kind],["Vendor",n.vendor||""],["MAC",n.id],["SSID",displaySsid(n.ssid,"")],["BSSID",n.bssid||""],["Degree",degree(id)],["Capture",(n.capture_files||[]).join("\n")],["Start samples",(n.start_samples||[]).join(", ")]])+
    `<div class="sectionTitle">Related Frames</div>${frameList(nodeFrames(id))}`;
  document.getElementById("openPlotsButton")?.addEventListener("click",()=>openPlots(id));renderList();draw();
}
function selectEdge(id){
  selectedType="edge";selectedId=id;const e=edgeMap.get(id);if(!e)return;const fs=graph.frames.filter(f=>(e.frames||[]).includes(f.index));setDetailHeader(`${e.source} -> ${e.target}`,"Link");details.className="panel";
  details.innerHTML=metrics([["Frames",e.weight||0],["Source degree",degree(e.source)],["Target degree",degree(e.target)]])+
    `<div class="sectionTitle">Link</div>`+rows([["Source",e.source],["Target",e.target],["SSID",displaySsid(e.ssid,"")],["Relation",e.relation||""]])+
    `<div class="sectionTitle">Frames</div>${frameList(fs.slice(0,30))}`;
  renderList();draw();
}
function renderStats(){
  stats.innerHTML=[["Nodes",summary.node_count||graph.nodes.length],["Edges",summary.edge_count||graph.edges.length],["Frames",summary.frame_count||graph.frames.length]].map(([k,v])=>`<div class="stat"><b>${esc(v)}</b><span>${esc(k)}</span></div>`).join("");
  const source=(summary.capture_files&&summary.capture_files[0])||graph.source_dir||"";sourcePath.textContent=source;stageSummary.textContent=`${graph.nodes.length} devices, ${graph.edges.length} links`;
  badges.innerHTML=[`Output: ${graph.source_xlsx||"output.xlsx"}`,`Capture files: ${(summary.capture_files||[]).length||1}`].map(x=>`<span class="badge">${esc(x)}</span>`).join("");
}
function renderLegend(){legend.innerHTML=["Access Point","Hidden AP","Station","Broadcast"].map(k=>`<span class="chip"><span class="dot" style="background:${colorKind(k)}"></span>${esc(k)}</span>`).join("");}
function renderList(){
  const q=filter.value.trim().toLowerCase();
  const list=graph.nodes.filter(n=>!q||[n.id,n.label,n.kind,n.vendor,n.ssid,n.bssid].join(" ").toLowerCase().includes(q)).sort((a,b)=>kindRank(a)-kindRank(b)||Number(b.frames||0)-Number(a.frames||0)||String(a.id).localeCompare(String(b.id)));
  nodeList.innerHTML=list.map(n=>`<button class="nodeItem ${selectedType==="node"&&selectedId===n.id?"active":""}" data-id="${esc(n.id)}"><span class="swatch" style="background:${colorKind(n.kind)}"></span><span><span class="name">${esc(n.label||n.id)}</span><span class="meta">${esc(n.vendor?`${n.vendor} ${n.kind}`:n.kind)} · ${esc(displaySsid(n.ssid,"no SSID"))}</span></span><span class="count">${esc(n.frames||0)}</span></button>`).join("")||`<div class="empty">No nodes.</div>`;
  nodeList.querySelectorAll(".nodeItem").forEach(el=>{el.addEventListener("click",()=>selectNode(el.dataset.id));el.addEventListener("dblclick",()=>openPlots(el.dataset.id));});
}
function overview(){
  selectedType="";selectedId="";setDetailHeader("Topology Overview","Overview");details.className="panel";
  const top=[...graph.nodes].sort((a,b)=>Number(b.frames||0)-Number(a.frames||0)).slice(0,8);
  details.innerHTML=metrics([["AP / Hidden AP",graph.nodes.filter(n=>n.kind==="Access Point"||n.kind==="Hidden AP").length],["Intel Stations",graph.nodes.filter(n=>n.vendor==="Intel").length],["Broadcast",graph.nodes.filter(n=>n.kind==="Broadcast").length]])+
    `<div class="sectionTitle">Most Active Nodes</div>`+top.map(n=>`<div class="frame"><b>${esc(n.label||n.id)}</b><br>${esc(n.vendor?`${n.vendor} ${n.kind}`:n.kind)} · ${esc(n.frames||0)} frames · degree ${degree(n.id)}</div>`).join("");
}
function setPlot(img,n,which){const fig=img.closest(".plotFigure");fig.classList.remove("missing");img.onload=()=>fig.classList.remove("missing");img.onerror=()=>fig.classList.add("missing");img.alt=`${n.id} ${which} plot`;img.src=plotPath(n,which);}
function openPlots(id){const n=nodeMap.get(id);if(!n)return;document.getElementById("modalTitle").textContent=n.label||n.id;document.getElementById("modalSub").textContent=`${n.kind} · ${n.frames||0} frames`;setPlot(document.getElementById("timelinePlot"),n,"timeline");setPlot(document.getElementById("framePlot"),n,"frame");document.getElementById("plotModal").classList.add("show");document.getElementById("plotModal").setAttribute("aria-hidden","false");}
function closePlots(){document.getElementById("plotModal").classList.remove("show");document.getElementById("plotModal").setAttribute("aria-hidden","true");}
document.getElementById("closeModal").addEventListener("click",closePlots);
document.getElementById("plotModal").addEventListener("click",evt=>{if(evt.target.id==="plotModal")closePlots();});
document.addEventListener("keydown",evt=>{if(evt.key==="Escape")closePlots();});
filter.addEventListener("input",renderList);
svg.addEventListener("wheel",evt=>{evt.preventDefault();zoomAt(evt.clientX,evt.clientY,evt.deltaY<0?1.14:0.88);},{passive:false});
svg.addEventListener("pointerdown",evt=>{
  if(evt.button!==0)return;
  if(evt.target.closest(".node,.apNode,.edge,.edgeHit"))return;
  dragState={id:evt.pointerId,x:evt.clientX,y:evt.clientY,startX:view.x,startY:view.y};
  svg.setPointerCapture(evt.pointerId);svg.classList.add("panning");
});
svg.addEventListener("pointermove",evt=>{
  if(nodeDrag&&evt.pointerId===nodeDrag.pointerId){
    const dx=(evt.clientX-nodeDrag.startClientX)/view.scale,dy=(evt.clientY-nodeDrag.startClientY)/view.scale;
    if(Math.abs(evt.clientX-nodeDrag.startClientX)>3||Math.abs(evt.clientY-nodeDrag.startClientY)>3)nodeDrag.moved=true;
    manualPos.set(nodeDrag.id,{x:nodeDrag.startX+dx,y:nodeDrag.startY+dy});
    draw();return;
  }
  if(!dragState||evt.pointerId!==dragState.id)return;
  view.x=dragState.startX+evt.clientX-dragState.x;view.y=dragState.startY+evt.clientY-dragState.y;applyView();
});
svg.addEventListener("pointerup",evt=>{
  if(nodeDrag&&evt.pointerId===nodeDrag.pointerId){if(nodeDrag.moved)suppressNodeClick=nodeDrag.id;nodeDrag=null;svg.classList.remove("panning");return;}
  if(dragState&&evt.pointerId===dragState.id){dragState=null;svg.classList.remove("panning");}
});
svg.addEventListener("pointercancel",()=>{dragState=null;nodeDrag=null;svg.classList.remove("panning");});
zoomInButton.addEventListener("click",()=>zoomAt(svg.getBoundingClientRect().left+svg.clientWidth/2,svg.getBoundingClientRect().top+svg.clientHeight/2,1.22));
zoomOutButton.addEventListener("click",()=>zoomAt(svg.getBoundingClientRect().left+svg.clientWidth/2,svg.getBoundingClientRect().top+svg.clientHeight/2,0.82));
zoomResetButton.addEventListener("click",resetView);
window.addEventListener("resize",draw);
renderStats();renderLegend();renderList();overview();draw();
</script>
</body>
</html>"""
    return template.replace("__GRAPH_JSON__", graph_json)

def topology_html_3d(graph: Dict[str, object]) -> str:
    graph_json = json.dumps(graph, ensure_ascii=False).replace("</", "<\\/")
    template = r"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>WiFi Signal Topology 3D</title>
<style>
:root{--bg:#eef2f6;--panel:#fff;--soft:#f8fafc;--line:#d8e0ea;--text:#172033;--muted:#667085;--blue:#2563eb;--green:#079669;--purple:#7c3aed;--red:#dc2626}
*{box-sizing:border-box}
body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--text);letter-spacing:0}
.app{display:grid;grid-template-columns:316px minmax(560px,1fr) 392px;height:100vh;min-height:640px}
.side,.details{background:var(--panel);overflow:auto}
.side{border-right:1px solid var(--line)}
.details{border-left:1px solid var(--line)}
.brand{padding:20px 18px 14px;border-bottom:1px solid var(--line)}
.brandTop{display:flex;align-items:flex-start;justify-content:space-between;gap:10px}
.brandTop h1{margin-bottom:0;min-width:0}
.brandControls{display:flex;flex-direction:column;align-items:flex-end;gap:6px;flex:0 0 auto}
.settingsButton{width:32px;height:32px;display:inline-flex;align-items:center;justify-content:center;border:1px solid #cbd5e1;border-radius:7px;background:#fff;color:#475467;font-size:17px;line-height:1;cursor:pointer;flex:0 0 auto}
.settingsButton:hover{background:#f8fafc;border-color:#94a3b8;color:#172033}
.langSwitch{display:flex;border:1px solid #cbd5e1;border-radius:7px;overflow:hidden;background:#fff}
.langOption{height:24px;border:0;border-right:1px solid #e2e8f0;background:#fff;color:#475467;font-size:11px;padding:0 7px;cursor:pointer;line-height:1}
.langOption:last-child{border-right:0}
.langOption:hover{background:#f8fafc}
.langOption.active{background:#172033;color:#fff}
h1{font-size:18px;line-height:1.2;margin:0 0 7px;font-weight:700}
.subtle{font-size:12px;color:var(--muted);line-height:1.45;word-break:break-all}
.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;padding:14px 14px 12px}
.stat{background:var(--soft);border:1px solid var(--line);border-radius:8px;padding:10px 9px;min-width:0}
.stat b{display:block;font-size:19px;line-height:1;color:var(--text)}
.stat span{display:block;margin-top:5px;font-size:11px;color:var(--muted);white-space:nowrap}
.tools{padding:0 14px 12px;border-bottom:1px solid var(--line)}
.search{width:100%;height:36px;padding:7px 10px;border:1px solid #cbd5e1;border-radius:7px;background:#fff;font-size:13px;outline:none}
.search:focus{border-color:#7aa7ff;box-shadow:0 0 0 3px #dbeafe}
.legend{display:flex;flex-wrap:wrap;gap:7px;margin-top:10px}
.chip{display:inline-flex;align-items:center;gap:6px;min-height:24px;padding:3px 8px;border:1px solid var(--line);border-radius:999px;background:#fff;font-size:11px;color:#475467}
.dot{width:8px;height:8px;border-radius:50%;display:inline-block}
.nodeList{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:8px;padding:8px 8px 16px;align-items:start}
.nodeColumn{min-width:0}
.nodeColumnHead{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:8px 7px 7px;color:#475467;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em}
.nodeColumnHead b{font-size:11px;color:#667085;background:#f1f5f9;border-radius:999px;padding:2px 6px}
.nodeItem{width:100%;display:grid;grid-template-columns:12px minmax(0,1fr) auto;gap:7px;align-items:center;padding:8px 7px;border:1px solid #edf1f6;border-radius:8px;background:#fff;color:inherit;text-align:left;cursor:pointer;margin-bottom:7px;min-width:0}
.nodeItem:hover,.nodeItem.active{background:#eef6ff;border-color:#bfdbfe}
.nodeItem.sourceTarget{border-color:#d4a017;box-shadow:0 0 0 1px #d4a017;background:#fffdf5}
.nodeItem.sourceTarget:hover,.nodeItem.sourceTarget.active{background:#fff7d6;border-color:#d4a017}
.nodeItem.sourceOnly{border-color:#111827;box-shadow:0 0 0 1px #111827}
.nodeItem.sourceOnly:hover,.nodeItem.sourceOnly.active{background:#f8fafc;border-color:#111827}
.nodeItem.targetOnly{border-color:var(--node-color,#079669);border-style:dashed;box-shadow:0 0 0 1px transparent;background:#fff}
.nodeItem.targetOnly:hover,.nodeItem.targetOnly.active{background:#f8fffb;border-color:var(--node-color,#079669)}
.nodeItem .name{display:block;font-size:12px;font-weight:650;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.nodeItem .meta{display:block;font-size:10.5px;color:var(--muted);margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.nodeItem .count{font-size:11px;font-variant-numeric:tabular-nums;color:#475467;background:#f1f5f9;border-radius:6px;padding:2px 5px}
.swatch{width:11px;height:11px;border-radius:50%;box-shadow:0 0 0 3px #fff,0 0 0 4px #d8e0ea}
.stage{position:relative;overflow:hidden;background:linear-gradient(180deg,#f8fbff 0%,#eef4f8 58%,#e8eef3 100%);user-select:none;-webkit-user-select:none}
.stage *{user-select:none;-webkit-user-select:none}
.stageTop{position:absolute;left:18px;right:18px;top:14px;z-index:2;display:flex;justify-content:space-between;align-items:flex-start;gap:12px;pointer-events:none}
.stageTitle{background:#fffffff2;border:1px solid var(--line);border-radius:8px;padding:10px 12px;box-shadow:0 8px 24px #34405414}
.stageTitle b{display:block;font-size:13px;margin-bottom:4px}
.stageTitle span{font-size:12px;color:var(--muted)}
.badgeRow{display:flex;flex-wrap:wrap;justify-content:flex-end;gap:7px}
.badge{background:#fffffff2;border:1px solid var(--line);border-radius:999px;padding:6px 9px;font-size:11px;color:#475467;box-shadow:0 8px 24px #34405412}
svg{width:100%;height:100%;display:block;touch-action:none;cursor:grab}
svg.panning{cursor:grabbing}
.zoomControls{position:absolute;right:18px;bottom:18px;z-index:3;display:flex;gap:6px;background:#fffffff2;border:1px solid var(--line);border-radius:8px;padding:6px;box-shadow:0 8px 24px #34405418}
.zoomButton{min-width:32px;height:32px;border:1px solid #cbd5e1;border-radius:7px;background:#fff;color:#172033;font-size:13px;font-weight:700;line-height:1;cursor:pointer;padding:0 9px}
.zoomButton:hover{background:#f8fafc;border-color:#94a3b8}
.hint{position:absolute;left:18px;bottom:18px;z-index:2;background:#fffffff0;border:1px solid var(--line);border-radius:8px;padding:8px 10px;color:#667085;font-size:12px;box-shadow:0 8px 24px #34405412;pointer-events:none}
.localTopology{position:absolute;left:18px;bottom:58px;z-index:3;width:286px;background:#fffffff2;border:1px solid var(--line);border-radius:8px;box-shadow:0 12px 34px #34405420;overflow:hidden;display:none}
.localTopology.show{display:block}
.localTop{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:9px 10px;border-bottom:1px solid var(--line);background:#fff}
.localTop b{font-size:12px;color:#172033}
.localTop span{font-size:11px;color:var(--muted);white-space:nowrap}
.localSvg{width:100%;height:220px;display:block;background:#f8fbff;cursor:default}
.localEdge{stroke:#64748b;stroke-width:1.8;stroke-linecap:round}
.localEdge.dashed{stroke-dasharray:5 4}
.localEdge.active{stroke:#2563eb;stroke-width:2.6}
.localEdgeLabel{font-size:9px;fill:#475467;text-anchor:middle;paint-order:stroke;stroke:#fff;stroke-width:3px;stroke-linejoin:round}
.localNode{stroke:#fff;stroke-width:2.5;cursor:pointer;filter:drop-shadow(0 4px 7px #0f172a24)}
.localNode.center{stroke:#111827;stroke-width:3}
.localRoleRing{fill:none;stroke:#111827;stroke-width:2.3;pointer-events:none}
.localRoleRing.sourceOnly{stroke:#111827}
.localRoleRing.sourceTarget{stroke:#d4a017;stroke-width:3}
.localRoleRing.targetOnly{stroke:var(--node-color,#079669);stroke-width:2.4;stroke-dasharray:5 4}
.localLabel{font-size:9px;fill:#172033;text-anchor:middle;font-weight:650;paint-order:stroke;stroke:#fff;stroke-width:3px;stroke-linejoin:round}
.localKind{font-size:8px;fill:#667085;text-anchor:middle;paint-order:stroke;stroke:#fff;stroke-width:3px;stroke-linejoin:round}
.groundLine{stroke:#cbd5e1;stroke-width:1;stroke-opacity:.52}
.axisLine{stroke:#94a3b8;stroke-width:1.4;stroke-opacity:.75}
.axisLabel{font-size:10px;fill:#667085;text-anchor:middle;pointer-events:none}
.edge{stroke:#7c8da6;stroke-opacity:.68;cursor:pointer;fill:none;stroke-linecap:round}
.edge.related{stroke:#4f7cff;stroke-opacity:.94}
.edge.selected{stroke:var(--red);stroke-opacity:1}
.edge.faded{stroke-opacity:.13}
.edge.dashed{stroke-dasharray:8 7}
.edgeHit{stroke:transparent;stroke-width:16;cursor:pointer;fill:none}
.node{stroke:#fff;stroke-width:2.5;cursor:pointer;filter:drop-shadow(0 8px 10px #0f172a28)}
.node.related{stroke:#172033;stroke-width:2.5}
.node.selected{stroke:#111827;stroke-width:4}
.node.faded{opacity:.28}
.node.targetOnly{fill:#fff;stroke:#fff;stroke-width:2.5;filter:none}
.nodeRoleRing{fill:none;stroke:#111827;stroke-width:3.2;pointer-events:none}
.nodeRoleRing.sourceOnly{stroke:#111827}
.nodeRoleRing.sourceTarget{stroke:#d4a017;stroke-width:4}
.nodeRoleRing.targetOnly{stroke:var(--node-color,#079669);stroke-width:2.8;stroke-dasharray:7 5}
.nodeHit{fill:#fff;fill-opacity:0;stroke:transparent;pointer-events:all;cursor:pointer}
.apNode,.iconNode{cursor:pointer}
.apNode.faded,.iconNode.faded{opacity:.28}
.routerImage,.deviceImage{pointer-events:auto;cursor:pointer;filter:drop-shadow(0 9px 10px #0f172a25)}
.apNode.selected .routerImage,.iconNode.selected .deviceImage{filter:drop-shadow(0 10px 13px #11182760)}
.targetOnlyRing{fill:none;stroke:var(--node-color,#2563eb);stroke-width:2.6;stroke-dasharray:7 5;pointer-events:none}
.label{font-size:11px;text-anchor:middle;fill:#1f2937;pointer-events:none;font-weight:650}
.kindLabel{font-size:9px;text-anchor:middle;fill:#667085;pointer-events:none}
.edgeLabel{font-size:11px;fill:#475467;text-anchor:middle;pointer-events:none;font-variant-numeric:tabular-nums}
.edgeLabelBg{fill:#fff;stroke:#d8e0ea;stroke-width:1}
.panel{padding:17px 16px 22px}
.panelHead{padding:18px 16px 14px;border-bottom:1px solid var(--line);background:#fff;position:sticky;top:0;z-index:3}
.panelHead h2{font-size:17px;line-height:1.25;margin:0 0 7px;word-break:break-all}
.typePill{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--line);border-radius:999px;padding:4px 9px;font-size:11px;color:#475467;background:#fff}
.metricGrid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin:14px 0}
.metric{background:var(--soft);border:1px solid var(--line);border-radius:8px;padding:9px}
.metric b{display:block;font-size:18px;line-height:1}
.metric span{display:block;margin-top:5px;font-size:11px;color:var(--muted)}
table{width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed}
td{border-bottom:1px solid #edf1f6;padding:8px 4px;vertical-align:top;word-break:break-word}
td:first-child{color:var(--muted);width:112px}
.sectionTitle{font-size:12px;text-transform:uppercase;color:#667085;letter-spacing:.04em;margin:18px 0 8px}
.frame{font-size:12px;border:1px solid #e7edf5;background:#fbfdff;border-radius:8px;padding:9px 10px;margin:7px 0;line-height:1.45}
.empty{padding:18px 16px;color:#667085;font-size:13px;line-height:1.5}
.actions{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0 4px}
.button{height:34px;border:1px solid #cbd5e1;border-radius:7px;background:#fff;color:#172033;padding:0 11px;font-size:12px;cursor:pointer}
.button:hover{background:#f8fafc;border-color:#94a3b8}
.modal{position:fixed;inset:0;background:#0f172a66;display:none;align-items:center;justify-content:center;padding:28px;z-index:10}
.modal.show{display:flex}
.modalShell{width:min(1180px,96vw);max-height:92vh;overflow-y:auto;overflow-x:hidden;background:#fff;border-radius:10px;box-shadow:0 24px 80px #0f172a66;border:1px solid #d8e0ea}
.modalTop{position:sticky;top:0;background:#fff;border-bottom:1px solid var(--line);padding:13px 15px;display:flex;justify-content:space-between;align-items:center;gap:12px;z-index:2}
.modalTop b{display:block;font-size:15px;word-break:break-all}
.modalTop span{font-size:12px;color:var(--muted)}
.plotGrid{display:grid;grid-template-columns:1fr 1fr;gap:14px;padding:14px}
.plotFigure{margin:0;border:1px solid #d8e0ea;border-radius:8px;background:#f8fafc;min-height:280px;overflow:hidden}
.plotFigure figcaption{font-size:12px;color:#475467;padding:10px 12px;border-bottom:1px solid #d8e0ea;background:#fff}
.plotFigure img{display:block;width:100%;height:auto;background:#fff}
.plotFigure .missingNote{display:none;margin:0;padding:28px 16px;color:#667085;font-size:13px}
.plotFigure.missing img{display:none}
.plotFigure.missing .missingNote{display:block}
.abbrTable{padding:14px}
.abbrTable table{font-size:13px}
.abbrTable th{font-size:12px;text-align:left;color:#475467;background:#f8fafc;border-bottom:1px solid #d8e0ea;padding:9px 8px}
.abbrTable td{padding:9px 8px}
.abbrTable td:first-child{width:88px;font-weight:700;color:#172033}
.channelWaveforms{border-top:1px solid #d8e0ea;padding:14px 16px 18px;background:#fbfdff;overflow:hidden}
.channelWaveforms h3{font-size:13px;margin:0 0 10px;color:#172033}
.channelWaveforms .channelNote{font-size:12px;color:#667085;margin:0 0 12px;line-height:1.45}
.channelGrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:10px;min-width:0}
.channelCard{margin:0;border:1px solid #d8e0ea;border-radius:8px;background:#fff;overflow:hidden;min-width:0;max-width:100%}
.channelCard figcaption{display:flex;justify-content:space-between;gap:8px;align-items:center;padding:7px 9px;border-bottom:1px solid #edf1f6;font-size:12px;color:#172033;min-width:0}
.channelCard figcaption b{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.channelCard figcaption span{font-size:11px;color:#667085;white-space:nowrap}
.channelCard img{display:block;width:100%;max-width:100%;height:auto;object-fit:contain;background:#fff}
.channelEmpty{font-size:12px;color:#667085;margin:0}
.frameLegend{border-top:1px solid #d8e0ea;padding:14px 16px 18px;background:#fff}
.frameLegend h3{font-size:13px;margin:0 0 10px;color:#172033}
.frameLegendGrid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}
.frameLegendItem{border:1px solid #e7edf5;border-radius:8px;background:#fbfdff;padding:10px;min-width:0}
.frameLegendSwatch{width:38px;height:26px;border-radius:7px;background:#fff;margin-bottom:8px}
.frameLegendSwatch.sourceOnly{border:2px solid #111827}
.frameLegendSwatch.sourceTarget{border:2px solid #d4a017}
.frameLegendSwatch.targetOnly{border:2px dashed #079669}
.frameLegendItem b{display:block;font-size:12px;color:#172033;margin-bottom:5px}
.frameLegendItem p{font-size:12px;line-height:1.45;color:#667085;margin:0}
@media (max-width:1500px){.app{grid-template-columns:300px minmax(0,1fr);grid-template-rows:720px 340px;height:auto;min-height:100vh}.side,.stage{height:720px}.details{grid-column:1/-1;height:340px;border-left:0;border-top:1px solid var(--line)}}
@media (max-width:760px){.app{display:block;height:auto}.side,.details{height:auto;max-height:none;border:0;border-bottom:1px solid var(--line)}.stage{height:620px}.plotGrid,.channelGrid,.frameLegendGrid{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="app">
<aside class="side">
  <div class="brand"><div class="brandTop"><h1>WiFi Signal Topology 3D</h1><div class="brandControls"><button class="settingsButton" id="openAbbrModal" title="Abbreviations">⚙</button><div class="langSwitch" aria-label="Language"><button class="langOption active" id="langZh" data-lang="zh" type="button">中文</button><button class="langOption" id="langEn" data-lang="en" type="button">English</button></div></div></div><div class="subtle" id="sourcePath"></div></div>
  <div class="stats" id="stats"></div>
  <div class="tools"><input class="search" id="filter" placeholder="Search MAC / SSID / type"><div class="legend" id="legend"></div></div>
  <div class="nodeList" id="nodeList"></div>
</aside>
<main class="stage">
  <div class="stageTop"><div class="stageTitle"><b>3D Decoded Device Graph</b><span id="stageSummary"></span></div><div class="badgeRow" id="badges"></div></div>
  <div class="zoomControls"><button class="zoomButton" id="zoomOut" title="Zoom out">-</button><button class="zoomButton" id="zoomReset" title="Reset 3D view">Reset</button><button class="zoomButton" id="zoomIn" title="Zoom in">+</button></div>
  <div class="localTopology" id="localTopology"><div class="localTop"><b>Local Topology</b><span id="localTopologyMeta"></span></div><svg class="localSvg" id="localTopologySvg"></svg></div>
  <div class="hint">Drag empty space to rotate. Arrow keys or WASD move the camera viewpoint. Shift moves faster. Drag a node to move it.</div>
  <svg id="graph"></svg>
</main>
<section class="details"><div class="panelHead"><h2>Topology Overview</h2><span class="typePill"><span class="dot" style="background:#94a3b8"></span>Overview</span></div><div id="details" class="empty"></div></section>
</div>
<div class="modal" id="plotModal" aria-hidden="true"><div class="modalShell"><div class="modalTop"><div><b id="modalTitle"></b><span id="modalSub"></span></div><button class="button" id="closeModal">Close</button></div><div class="plotGrid"><figure class="plotFigure"><figcaption>Extracted Signal Timeline</figcaption><img id="timelinePlot" alt=""><p class="missingNote">No timeline image found in signal_plots.</p></figure><figure class="plotFigure"><figcaption>Single Frame Detail</figcaption><img id="framePlot" alt=""><p class="missingNote">No frame image found in signal_plots.</p></figure></div></div></div>
<div class="modal" id="abbrModal" aria-hidden="true"><div class="modalShell"><div class="modalTop"><div><b>Abbreviations</b><span>English full name and Chinese translation</span></div><button class="button" id="closeAbbrModal">Close</button></div><div class="abbrTable" id="abbrTable"></div><div class="abbrReferences"><h3>相关文献：</h3><p>Z. Zhou, C. Hou, Z. Song, B. Wang, X. Wu and Z. Liu, &quot;Wireless Communication Network Multi-Node Topology Inference Method With Maximizing Information Utilization Based on GCN,&quot; in IEEE Transactions on Cognitive Communications and Networking, vol. 12, pp. 4319-4331, 2026.</p><p>Zhou Zhichao, Hou Changbo, Meng Guojing, et al. Research on the digital twin of communication relationships in the electromagnetic spectrum of wireless local area networks[J]. Journal on Communications, 2026, 47(05):78-90.</p><p class="author">哈尔滨工程大学周志超博士</p></div><div class="channelWaveforms"><h3>WiFi 子信道大尺度时域波形图</h3><p class="channelNote">每张图对应一个 20 Msps、1 s 采集文件，显示整段采集的归一化 |IQ| 峰值与均值包络。</p><div class="channelGrid" id="channelWaveformGrid"></div></div><div class="frameLegend" id="frameLegend"></div></div></div>
<script id="graph-data" type="application/json">__GRAPH_JSON__</script>
<script>
const graph=JSON.parse(document.getElementById("graph-data").textContent);
const svg=document.getElementById("graph"),details=document.getElementById("details"),stats=document.getElementById("stats"),nodeList=document.getElementById("nodeList"),filter=document.getElementById("filter"),legend=document.getElementById("legend"),badges=document.getElementById("badges"),sourcePath=document.getElementById("sourcePath"),stageSummary=document.getElementById("stageSummary"),localTopology=document.getElementById("localTopology"),localTopologySvg=document.getElementById("localTopologySvg"),localTopologyMeta=document.getElementById("localTopologyMeta");
let selectedType="",selectedId="",uiLang="zh",view={angleX:0.46,angleY:0.72,zoom:1,cameraX:0,cameraZ:0},dragState=null,nodeDrag=null,suppressNodeClick="",lastNodeClick={id:"",time:0},lastListClick={id:"",time:0},pendingNodeClickTimer=0,pendingListClickTimer=0;
const manualWorld=new Map(),routerIconPath="assets/router_2p5d_clean_node.png",intelLaptopIconPath="assets/laptop_2p5d_intel_node.png",nodeMap=new Map(graph.nodes.map(n=>[n.id,n])),edgeMap=new Map(graph.edges.map(e=>[e.id,e])),summary=graph.summary||{},adjacency=new Map();
const abbrItems=[
  ["AP","Access Point","无线接入点"],
  ["SSID","Service Set Identifier","WiFi 名称"],
  ["BSSID","Basic Service Set Identifier","无线接入点 MAC 地址"],
  ["MAC","Media Access Control","媒体访问控制地址"],
  ["FCS","Frame Check Sequence","帧校验序列"],
  ["MCS","Modulation and Coding Scheme","调制与编码方案"],
  ["IQ","In-phase and Quadrature","同相/正交采样信号"],
  ["USRP","Universal Software Radio Peripheral","通用软件无线电外设"],
  ["RX","Receive","接收"],
  ["TX","Transmit","发送"],
  ["3D","Three-dimensional","三维"],
  ["DFS","Dynamic Frequency Selection","动态频率选择"]
];
for(const n of graph.nodes)adjacency.set(n.id,new Set());
for(const e of graph.edges){if(!adjacency.has(e.source))adjacency.set(e.source,new Set());if(!adjacency.has(e.target))adjacency.set(e.target,new Set());adjacency.get(e.source).add(e.target);adjacency.get(e.target).add(e.source);}
const zhText={
  "Access Point":"接入点","Hidden AP":"隐藏接入点","Station":"工作站","Broadcast":"广播","Link":"链路","Overview":"总览","Selection":"选择",
  "Nodes":"节点","Edges":"连线","Frames":"帧","Sent":"发送","Received":"接收","Device":"设备","Type":"类型","Vendor":"厂商","Degree":"度","Capture":"采集文件","Start samples":"起始采样点","Related Frames":"相关帧","Source":"源节点","Target":"目标节点","Source degree":"源节点度","Target degree":"目标节点度","Relation":"关系","Line style":"线型","Dashed":"虚线","Solid":"实线",
  "Open plots":"打开信号图","Close":"关闭","Topology Overview":"拓扑总览","Most Active Nodes":"最活跃节点","AP / Hidden AP":"接入点/隐藏接入点","Intel Stations":"Intel 工作站","Local Topology":"局部拓扑","links":"条连接","no SSID":"无 WiFi 名称","Hidden WiFi":"隐藏 WiFi","No nodes.":"无节点","No frames.":"无帧","AP":"接入点","3D Decoded Device Graph":"3D 解码设备图","devices":"个设备","Output":"输出","Capture files":"采集文件","3D view":"3D 视图","Search MAC / SSID / type":"搜索 MAC / WiFi 名称 / 类型","Zoom out":"缩小","Reset":"重置","Reset 3D view":"重置 3D 视角","Zoom in":"放大","Drag empty space to rotate. Arrow keys or WASD move the camera viewpoint. Shift moves faster. Drag a node to move it.":"拖动空白处旋转视角。方向键或 WASD 移动视角，按住 Shift 加速。拖动节点可移动节点。","Extracted Signal Timeline":"提取信号大尺度时域图","Single Frame Detail":"单帧特写","No timeline image found in signal_plots.":"signal_plots 中未找到大尺度时域图。","No frame image found in signal_plots.":"signal_plots 中未找到单帧特写图。","Abbreviations":"英文简称","English full name and Chinese translation":"英文全称与中文翻译","Abbr.":"简称","Full Name":"英文全称","Chinese Translation":"中文翻译","No channel waveform images found. Regenerate topology with plot generation enabled.":"未找到子信道波形图，请开启绘图后重新生成拓扑。"
  ,"Node Frame Meaning":"节点线框含义","Source only":"仅源地址","Source and target":"源地址和目标地址","Target only":"仅目标地址","Black solid frame: the node appears only as a source address, meaning radiated signal is detected from this device.":"黑色实线框：节点只在源地址中出现，表示检测到该设备自身发射的信号。","Gold solid frame: the node appears as both source and target address.":"金色实线框：节点既在源地址中出现，也在目标地址中出现。","Dashed frame: the node is identified only from target address, without corresponding radiated signal.":"虚线框：节点只从目标地址中识别出来，没有对应的辐射信号。"
};
function tr(v){const s=String(v??"");return uiLang==="zh"?(zhText[s]||s):s;}
function kindText(kind){return tr(kind||"Selection");}
function esc(v){return String(v??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]));}
function makeSvg(t){return document.createElementNS("http://www.w3.org/2000/svg",t);}
function clamp(v,min,max){return Math.max(min,Math.min(max,v));}
function colorKind(k){return k==="Access Point"?"#2563eb":k==="Hidden AP"?"#7c3aed":k==="Broadcast"?"#dc2626":k==="Link"?"#94a3b8":"#079669";}
function shortMac(id){return id==="ff:ff:ff:ff:ff:ff"?"Broadcast":String(id||"").slice(-8);}
function cleanDisplayText(value){const raw=String(value??"");if(raw.trim()==="<hidden>")return "";let s=raw.replace(/<hidden>/gi," ").replace(/[\u0000-\u001f\u007f-\u009f\ufffd]/g," ");s=Array.from(s).filter(ch=>/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}A-Za-z0-9\s\-_.:@#[\]()+=~!$%^&,]/u.test(ch)).join("").replace(/\s+/g," ").trim();return Array.from(s).some(ch=>/[\p{L}\p{N}]/u.test(ch))?s:"";}
function displaySsid(value,fallback=""){const raw=String(value??""),clean=cleanDisplayText(raw);return clean||((raw.trim()==="<hidden>")?tr("Hidden WiFi"):fallback);}
function ellipsize(v,m){const s=String(v??"");return s.length>m?s.slice(0,m-3)+"...":s;}
function apTitle(n){return displaySsid(n.ssid,shortMac(n.id));}
function apLabel(n){return ellipsize(apTitle(n),26);}
function isApNode(n){return n&&(n.kind==="Access Point"||n.kind==="Hidden AP");}
function isTargetOnlyNode(n){return !!n&&(n.target_only===true||n.kind==="Broadcast"||((Number(n.sent)||0)===0&&(Number(n.received)||0)>0));}
function nodeAddressRole(n){if(!n)return"";const sent=Number(n.sent)||0,received=Number(n.received)||0;if(n.kind==="Broadcast"||sent===0&&received>0)return"targetOnly";if(sent>0&&received>0)return"sourceTarget";if(sent>0)return"sourceOnly";return"";}
function nodeFillColor(n){const role=nodeAddressRole(n);return(role==="sourceOnly"||role==="sourceTarget")&&!isApNode(n)&&!usesDeviceIcon(n)?"#004CA1":isTargetOnlyNode(n)?"#fff":colorKind(n.kind);}
function isRadiatedNode(n){return !!n&&n.has_radiated_signal===true&&!isTargetOnlyNode(n);}
function edgeUsesDashedLine(e){return !(isRadiatedNode(nodeMap.get(e.source))&&isRadiatedNode(nodeMap.get(e.target)));}
function degree(id){return adjacency.get(id)?.size||0;}
function kindRank(n){return n.kind==="Access Point"?0:n.kind==="Hidden AP"?1:n.kind==="Station"?2:3;}
function nodeSort(a,b){return Number(b.frames||0)-Number(a.frames||0)||String(a.id).localeCompare(String(b.id));}
function usesDeviceIcon(n){return n&&n.device_icon==="intel_laptop"&&!isApNode(n)&&n.kind!=="Broadcast";}
function apIconSize(n){return Math.max(58,Math.min(82,54+Math.sqrt(Number(n.frames)||1)*3));}
function deviceIconSize(n){return Math.max(54,Math.min(78,50+Math.sqrt(Number(n.frames)||1)*2.8));}
function nodeRadius(n){return isApNode(n)?apIconSize(n)*0.48:(usesDeviceIcon(n)?deviceIconSize(n)*0.44:Math.max(13,Math.min(32,10+Math.sqrt(Number(n.frames)||1)*2.6)));}
function plotStem(id){return String(id||"node").toLowerCase().replace(/[^a-z0-9]+/g,"_").replace(/^_+|_+$/g,"")||"node";}
function plotPath(n,which){const key=which==="timeline"?"timeline_plot":"frame_plot";return n[key]||`signal_plots/${plotStem(n.id)}_${which}.png`;}
function selectedNodeRelated(id){return selectedType!=="node"||id===selectedId||adjacency.get(selectedId)?.has(id);}
function edgeRelated(e){return !selectedType||(selectedType==="edge"&&e.id===selectedId)||(selectedType==="node"&&(e.source===selectedId||e.target===selectedId));}
function connectedComponents(){const seen=new Set(),out=[];for(const n of [...graph.nodes].sort(nodeSort)){if(seen.has(n.id))continue;const stack=[n.id],ids=[];seen.add(n.id);while(stack.length){const id=stack.pop();ids.push(id);for(const next of adjacency.get(id)||[])if(!seen.has(next)){seen.add(next);stack.push(next);}}out.push(ids.map(id=>nodeMap.get(id)).filter(Boolean).sort((a,b)=>kindRank(a)-kindRank(b)||nodeSort(a,b)));}return out.sort((a,b)=>b.length-a.length||String(a[0]?.id||"").localeCompare(String(b[0]?.id||"")));}
function componentEdges(nodes){const ids=new Set(nodes.map(n=>n.id));return graph.edges.filter(e=>ids.has(e.source)&&ids.has(e.target));}
function primaryAp(id,edges){let best=null,w=-1;for(const e of edges){if(e.source!==id&&e.target!==id)continue;const other=nodeMap.get(e.source===id?e.target:e.source);if(!isApNode(other))continue;if(Number(e.weight)>w){best=other;w=Number(e.weight);}}return best;}
function placeSphere(c,r,i,count,phase=0){if(count<=1)return{x:c.x+r,y:c.y,z:c.z};const golden=Math.PI*(3-Math.sqrt(5)),y=1-(i/(count-1))*2,ring=Math.sqrt(Math.max(0,1-y*y)),theta=i*golden+phase;return{x:c.x+Math.cos(theta)*ring*r,y:c.y+y*r*.78,z:c.z+Math.sin(theta)*ring*r};}
function gridCenters(count){const cols=Math.ceil(Math.sqrt(count)),spacing=620,centers=[];for(let i=0;i<count;i++){const row=Math.floor(i/cols),col=i%cols;centers.push({x:(col-(cols-1)/2)*spacing,y:0,z:(row-(Math.ceil(count/cols)-1)/2)*spacing});}return centers;}
function buildWorldLayout(){const comps=connectedComponents(),centers=gridCenters(comps.length),world=new Map();comps.forEach((nodes,ci)=>{const c=centers[ci]||{x:0,y:0,z:0},edges=componentEdges(nodes),aps=nodes.filter(isApNode).sort(nodeSort),bcasts=nodes.filter(n=>n.kind==="Broadcast").sort(nodeSort),stations=nodes.filter(n=>!isApNode(n)&&n.kind!=="Broadcast").sort(nodeSort),anchors=aps.length?aps:[...nodes].sort((a,b)=>degree(b.id)-degree(a.id)||kindRank(a)-kindRank(b)||nodeSort(a,b)).slice(0,1),inner=anchors.length>1?Math.max(92,Math.min(160,62+anchors.length*16)):0;anchors.forEach((n,i)=>{const a=-Math.PI/2+i*Math.PI*2/Math.max(anchors.length,1);world.set(n.id,anchors.length===1?{x:c.x,y:c.y,z:c.z}:{x:c.x+Math.cos(a)*inner,y:c.y+Math.sin(a*2)*28,z:c.z+Math.sin(a)*inner});});const groups=new Map(anchors.map(n=>[n.id,[]])),free=[];for(const n of stations){const ap=primaryAp(n.id,edges);if(ap&&groups.has(ap.id))groups.get(ap.id).push(n);else if(anchors.length===1)groups.get(anchors[0].id).push(n);else free.push(n);}for(const [anchorId,list] of groups.entries()){const anchor=world.get(anchorId)||c,r=Math.max(150,Math.min(265,138+Math.sqrt(Math.max(list.length,1))*28));list.sort((a,b)=>degree(b.id)-degree(a.id)||nodeSort(a,b)).forEach((n,i)=>world.set(n.id,placeSphere(anchor,r,i,list.length,ci*.73)));}const outer=Math.max(190,Math.min(320,170+Math.sqrt(nodes.length)*28));free.forEach((n,i)=>world.set(n.id,placeSphere(c,outer,i,free.length,Math.PI/3+ci)));bcasts.forEach((n,i)=>{const a=Math.PI*.75+(i-(bcasts.length-1)/2)*.32;world.set(n.id,{x:c.x+Math.cos(a)*outer*1.08,y:c.y-70,z:c.z+Math.sin(a)*outer*1.08});});if(nodes.length===1&&!world.has(nodes[0].id))world.set(nodes[0].id,{x:c.x,y:c.y,z:c.z});});for(const [id,p]of manualWorld.entries())world.set(id,{...p});return world;}
function spreadBroadcastNeighbors(world){
  for(const b of graph.nodes.filter(n=>n.kind==="Broadcast")){
    const bp=world.get(b.id);if(!bp)continue;
    const ids=[...new Set(graph.edges.filter(e=>e.source===b.id||e.target===b.id).map(e=>e.source===b.id?e.target:e.source))].filter(id=>world.has(id));
    if(ids.length<2)continue;
    const directIds=new Set(ids);
    const neighbors=ids.map(id=>nodeMap.get(id)).filter(Boolean).sort((a,c)=>kindRank(a)-kindRank(c)||nodeSort(a,c));
    const oldPos=neighbors.map(n=>world.get(n.id)).filter(Boolean);
    const cx=oldPos.reduce((s,p)=>s+p.x,0)/oldPos.length,cy=oldPos.reduce((s,p)=>s+p.y,0)/oldPos.length,cz=oldPos.reduce((s,p)=>s+p.z,0)/oldPos.length;
    const radius=Math.max(105,Math.min(210,88+neighbors.length*8));
    if(!manualWorld.has(b.id)){bp.x=cx;bp.y=cy+12;bp.z=cz;world.set(b.id,bp);}
    neighbors.forEach((n,i)=>{
      if(manualWorld.has(n.id))return;
      const old=world.get(n.id);if(!old)return;
      const count=neighbors.length,angle=-Math.PI/2+i*Math.PI*2/count,ring=radius*(i%2?1.08:.9),level=((i%3)-1)*36;
      const next={x:bp.x+Math.cos(angle)*ring,y:bp.y+level,z:bp.z+Math.sin(angle)*ring};
      const delta={x:next.x-old.x,y:next.y-old.y,z:next.z-old.z};
      world.set(n.id,next);
      for(const m of graph.nodes){
        if(m.id===n.id||m.kind==="Broadcast"||directIds.has(m.id)||manualWorld.has(m.id)||!world.has(m.id))continue;
        const ap=primaryAp(m.id,graph.edges);
        if(ap&&ap.id===n.id){
          const p=world.get(m.id);
          world.set(m.id,{x:p.x+delta.x,y:p.y+delta.y,z:p.z+delta.z});
        }
      }
    });
  }
  return world;
}
function rotatePoint(p){const cy=Math.cos(view.angleY),sy=Math.sin(view.angleY),cx=Math.cos(view.angleX),sx=Math.sin(view.angleX),px=p.x-(view.cameraX||0),pz=p.z-(view.cameraZ||0),x1=px*cy-pz*sy,z1=px*sy+pz*cy,y1=p.y;return{x:x1,y:y1*cx-z1*sx,z:y1*sx+z1*cx};}
function project(p){const rect=svg.getBoundingClientRect(),w=rect.width||900,h=rect.height||650,r=rotatePoint(p),d=1250,scale=view.zoom*d/(d+r.z);return{x:w/2+r.x*scale,y:h/2+r.y*scale,scale,z:r.z,world:p};}
function projectAll(world){const out=new Map();for(const [id,p]of world.entries())out.set(id,project(p));return out;}
function screenBasis(){const cy=Math.cos(view.angleY),sy=Math.sin(view.angleY),cx=Math.cos(view.angleX),sx=Math.sin(view.angleX);return{right:{x:cy,y:0,z:-sy},up:{x:-sx*sy,y:cx,z:-sx*cy}};}
function horizontalBasis(){const cy=Math.cos(view.angleY),sy=Math.sin(view.angleY);return{right:{x:cy,z:-sy},forward:{x:sy,z:cy}};}
function drawGround(){const size=2400,step=180,y=120;for(let x=-size;x<=size;x+=step){drawLine(project({x,y,z:-size}),project({x,y,z:size}),"groundLine");}for(let z=-size;z<=size;z+=step){drawLine(project({x:-size,y,z}),project({x:size,y,z}),"groundLine");}drawLine(project({x:-size,y,z:0}),project({x:size,y,z:0}),"axisLine");drawLine(project({x:0,y,z:-size}),project({x:0,y,z:size}),"axisLine");}
function drawLine(a,b,cls){const line=makeSvg("line");line.setAttribute("x1",a.x);line.setAttribute("y1",a.y);line.setAttribute("x2",b.x);line.setAttribute("y2",b.y);line.setAttribute("class",cls);svg.appendChild(line);}
function draw(){const world=spreadBroadcastNeighbors(buildWorldLayout()),proj=projectAll(world);svg.innerHTML="";drawGround();[...graph.edges].map(e=>({e,z:((proj.get(e.source)?.z||0)+(proj.get(e.target)?.z||0))/2})).sort((a,b)=>b.z-a.z).forEach(({e})=>{const a=proj.get(e.source),b=proj.get(e.target);if(!a||!b)return;const d=`M ${a.x} ${a.y} L ${b.x} ${b.y}`,hit=makeSvg("path");hit.setAttribute("d",d);hit.setAttribute("class","edgeHit");hit.addEventListener("click",()=>selectEdge(e.id));svg.appendChild(hit);const line=makeSvg("path");line.setAttribute("d",d);line.setAttribute("stroke-width",Math.max(1.4,Math.min(7,1.1+Math.sqrt(Number(e.weight)||1)*Math.max(.72,(a.scale+b.scale)/2))));const classes=["edge"];if(edgeUsesDashedLine(e))classes.push("dashed");if(selectedType==="edge"&&selectedId===e.id)classes.push("selected");else if(edgeRelated(e))classes.push("related");else classes.push("faded");line.setAttribute("class",classes.join(" "));line.addEventListener("click",()=>selectEdge(e.id));svg.appendChild(line);if((Number(e.weight)||0)>1||selectedId===e.id){const mx=(a.x+b.x)/2,my=(a.y+b.y)/2-6,bg=makeSvg("rect"),t=makeSvg("text");bg.setAttribute("x",mx-12);bg.setAttribute("y",my-12);bg.setAttribute("width",24);bg.setAttribute("height",16);bg.setAttribute("rx",5);bg.setAttribute("class","edgeLabelBg");t.setAttribute("x",mx);t.setAttribute("y",my);t.setAttribute("class","edgeLabel");t.textContent=e.weight;svg.appendChild(bg);svg.appendChild(t);}});[...graph.nodes].map(n=>({n,p:proj.get(n.id)})).filter(x=>x.p).sort((a,b)=>b.p.z-a.p.z).forEach(({n,p})=>drawNode(n,p,world.get(n.id)));} 
function addNodeHit(parent,id,wp,x,y,r){const hit=makeSvg("circle");hit.setAttribute("cx",x);hit.setAttribute("cy",y);hit.setAttribute("r",Math.max(24,r));hit.setAttribute("class","nodeHit");hit.addEventListener("pointerdown",evt=>startNodeDrag(evt,id,wp));parent.appendChild(hit);}
function addNodeRoleRing(parent,n,x,y,r,local=false,skipSourceOnly=false){const role=nodeAddressRole(n);if(!role||(skipSourceOnly&&role==="sourceOnly"))return;const ring=makeSvg("circle");ring.setAttribute("cx",x);ring.setAttribute("cy",y);ring.setAttribute("r",r);ring.setAttribute("class",`${local?"localRoleRing":"nodeRoleRing"} ${role}`);ring.style.setProperty("--node-color",colorKind(n.kind));parent.appendChild(ring);}
function drawNode(n,p,wp){const scale=clamp(p.scale,.45,1.65);if(isApNode(n)||usesDeviceIcon(n)){const isAp=isApNode(n),size=(isAp?apIconSize(n):deviceIconSize(n))*scale,g=makeSvg("g"),img=makeSvg("image"),classes=[isAp?"apNode":"iconNode"];if(selectedType==="node"&&selectedId===n.id)classes.push("selected");else if(selectedNodeRelated(n.id))classes.push("related");else classes.push("faded");if(isTargetOnlyNode(n))classes.push("targetOnly");g.setAttribute("class",classes.join(" "));g.style.setProperty("--node-color",colorKind(n.kind));g.addEventListener("pointerdown",evt=>startNodeDrag(evt,n.id,wp));addNodeRoleRing(g,n,p.x,p.y,size*.64,false,true);img.setAttribute("href",isAp?routerIconPath:intelLaptopIconPath);img.setAttribute("x",p.x-size*(isAp?0.66:0.68));img.setAttribute("y",p.y-size*(isAp?0.48:0.5));img.setAttribute("width",size*(isAp?1.32:1.36));img.setAttribute("height",size*(isAp?0.9:1));img.setAttribute("class",isAp?"routerImage":"deviceImage");g.appendChild(img);const title=makeSvg("title");title.textContent=`${isAp?apTitle(n):(n.label||n.id)} | ${kindText(n.kind)} | ${n.frames||0} ${tr("Frames")}`;g.appendChild(title);addNodeHit(g,n.id,wp,p.x,p.y,size*.68);svg.appendChild(g);drawLabel(n,p,size*(isAp?0.58:0.6),isAp?apLabel(n):shortMac(n.id),isAp?(n.kind==="Hidden AP"?kindText("Hidden AP"):tr("AP")):`Intel ${kindText("Station")}`);return;}const r=nodeRadius(n)*scale,c=makeSvg("circle"),nodeColor=colorKind(n.kind),targetOnly=isTargetOnlyNode(n);c.setAttribute("cx",p.x);c.setAttribute("cy",p.y);c.setAttribute("r",r);c.setAttribute("fill",nodeFillColor(n));const classes=["node"];if(selectedType==="node"&&selectedId===n.id)classes.push("selected");else if(selectedNodeRelated(n.id))classes.push("related");else classes.push("faded");if(targetOnly)classes.push("targetOnly");c.setAttribute("class",classes.join(" "));c.style.setProperty("--node-color",nodeColor);c.addEventListener("pointerdown",evt=>startNodeDrag(evt,n.id,wp));const title=makeSvg("title");title.textContent=`${n.label||n.id} | ${kindText(n.kind)} | ${n.frames||0} ${tr("Frames")}`;c.appendChild(title);svg.appendChild(c);addNodeRoleRing(svg,n,p.x,p.y,r+5);addNodeHit(svg,n.id,wp,p.x,p.y,r+10);drawLabel(n,p,r+15,shortMac(n.id),n.kind==="Broadcast"?kindText("Broadcast"):kindText("Station"),13);}
function drawLabel(n,p,yOff,text,kind,kindOff=15){const label=makeSvg("text");label.setAttribute("x",p.x);label.setAttribute("y",p.y+yOff);label.setAttribute("class","label");label.textContent=text;svg.appendChild(label);const k=makeSvg("text");k.setAttribute("x",p.x);k.setAttribute("y",p.y+yOff+kindOff);k.setAttribute("class","kindLabel");k.textContent=kind;svg.appendChild(k);}
function localNodeTitle(n){if(!n)return"";return isApNode(n)?apLabel(n):(n.kind==="Broadcast"?kindText("Broadcast"):shortMac(n.id));}
function localNodeKind(n){if(!n)return"";return n.kind==="Access Point"?tr("AP"):n.kind==="Hidden AP"?kindText("Hidden AP"):n.kind==="Broadcast"?kindText("Broadcast"):kindText("Station");}
function drawLocalNode(n,x,y,center=false){const r=center?18:13,c=makeSvg("circle");c.setAttribute("cx",x);c.setAttribute("cy",y);c.setAttribute("r",r);c.setAttribute("fill",nodeFillColor(n));const cls=["localNode"];if(center)cls.push("center");c.setAttribute("class",cls.join(" "));c.style.setProperty("--node-color",colorKind(n.kind));c.addEventListener("click",()=>selectNode(n.id));localTopologySvg.appendChild(c);addNodeRoleRing(localTopologySvg,n,x,y,r+3,true);const title=makeSvg("title");title.textContent=`${localNodeTitle(n)} | ${kindText(n.kind)} | ${n.frames||0} ${tr("Frames")}`;c.appendChild(title);const label=makeSvg("text");label.setAttribute("x",x);label.setAttribute("y",y+r+13);label.setAttribute("class","localLabel");label.textContent=ellipsize(localNodeTitle(n),18);localTopologySvg.appendChild(label);const kind=makeSvg("text");kind.setAttribute("x",x);kind.setAttribute("y",y+r+24);kind.setAttribute("class","localKind");kind.textContent=localNodeKind(n);localTopologySvg.appendChild(kind);}
function renderLocalTopology(id){const center=nodeMap.get(id);if(!center){localTopology.classList.remove("show");return;}const edges=graph.edges.filter(e=>e.source===id||e.target===id).sort((a,b)=>Number(b.weight||0)-Number(a.weight||0));const neighborIds=[...new Set(edges.map(e=>e.source===id?e.target:e.source))];const neighbors=neighborIds.map(x=>nodeMap.get(x)).filter(Boolean).sort((a,b)=>kindRank(a)-kindRank(b)||nodeSort(a,b));localTopology.classList.add("show");localTopologyMeta.textContent=`${neighbors.length} ${tr("links")}`;localTopologySvg.innerHTML="";localTopologySvg.setAttribute("viewBox","0 0 286 220");const cx=143,cy=95,radius=68,pos=new Map([[id,{x:cx,y:cy}]]);neighbors.forEach((n,i)=>{const count=Math.max(neighbors.length,1),angle=-Math.PI/2+i*Math.PI*2/count,rx=neighbors.length===1?0:Math.cos(angle)*radius,ry=neighbors.length===1?64:Math.sin(angle)*radius*.78;pos.set(n.id,{x:cx+rx,y:cy+ry});});edges.forEach(e=>{const a=pos.get(e.source),b=pos.get(e.target);if(!a||!b)return;const line=makeSvg("line");line.setAttribute("x1",a.x);line.setAttribute("y1",a.y);line.setAttribute("x2",b.x);line.setAttribute("y2",b.y);const cls=["localEdge"];if(edgeUsesDashedLine(e))cls.push("dashed");if(e.source===id||e.target===id)cls.push("active");line.setAttribute("class",cls.join(" "));localTopologySvg.appendChild(line);if(Number(e.weight||0)>1){const t=makeSvg("text");t.setAttribute("x",(a.x+b.x)/2);t.setAttribute("y",(a.y+b.y)/2-4);t.setAttribute("class","localEdgeLabel");t.textContent=e.weight;localTopologySvg.appendChild(t);}});neighbors.forEach(n=>{const p=pos.get(n.id);drawLocalNode(n,p.x,p.y,false);});drawLocalNode(center,cx,cy,true);}
function rows(items){return `<table>${items.map(([k,v])=>`<tr><td>${esc(k)}</td><td>${esc(v)}</td></tr>`).join("")}</table>`;}
function metrics(items){return `<div class="metricGrid">${items.map(([k,v])=>`<div class="metric"><b>${esc(v)}</b><span>${esc(k)}</span></div>`).join("")}</div>`;}
function nodeFrames(id){return graph.frames.filter(f=>[f.source,f.target,String(f.bssid||"").toLowerCase(),String(f.receiver||"").toLowerCase(),String(f.transmitter||"").toLowerCase()].includes(id)).slice(0,22);}
function frameList(frames){return frames.map(f=>`<div class="frame"><b>#${esc(f.index)}</b> ${esc(f.frame_type)}/${esc(f.frame_subtype)}<br>${esc(f.source||"unknown")} -> ${esc(f.target||"unknown")}<br>start=${esc(f.start_sample)}, FCS=${esc(f.fcs_ok)}, MCS=${esc(f.mcs)}</div>`).join("")||`<div class="empty">${esc(tr("No frames."))}</div>`;}
function setDetailHeader(title,kind){document.querySelector(".panelHead").innerHTML=`<h2>${esc(title)}</h2><span class="typePill"><span class="dot" style="background:${colorKind(kind)}"></span>${esc(kindText(kind))}</span>`;}
function selectNode(id){selectedType="node";selectedId=id;const n=nodeMap.get(id);if(!n)return;setDetailHeader(n.label||n.id,n.kind);details.className="panel";details.innerHTML=metrics([[tr("Frames"),n.frames||0],[tr("Sent"),n.sent||0],[tr("Received"),n.received||0]])+`<div class="actions"><button class="button" id="openPlotsButton">${esc(tr("Open plots"))}</button></div><div class="sectionTitle">${esc(tr("Device"))}</div>`+rows([[tr("Type"),kindText(n.kind)],[tr("Vendor"),n.vendor||""],["MAC",n.id],["SSID",displaySsid(n.ssid,"")],["BSSID",n.bssid||""],[tr("Degree"),degree(id)],[tr("Capture"),(n.capture_files||[]).join("\n")],[tr("Start samples"),(n.start_samples||[]).join(", ")]])+`<div class="sectionTitle">${esc(tr("Related Frames"))}</div>${frameList(nodeFrames(id))}`;document.getElementById("openPlotsButton")?.addEventListener("click",()=>openPlots(id));renderLocalTopology(id);renderList();draw();}
function selectEdge(id){selectedType="edge";selectedId=id;const e=edgeMap.get(id);if(!e)return;const fs=graph.frames.filter(f=>(e.frames||[]).includes(f.index));setDetailHeader(`${e.source} -> ${e.target}`,"Link");details.className="panel";details.innerHTML=metrics([[tr("Frames"),e.weight||0],[tr("Source degree"),degree(e.source)],[tr("Target degree"),degree(e.target)]])+`<div class="sectionTitle">${esc(tr("Link"))}</div>`+rows([[tr("Source"),e.source],[tr("Target"),e.target],["SSID",displaySsid(e.ssid,"")],[tr("Relation"),e.relation||""],[tr("Line style"),edgeUsesDashedLine(e)?tr("Dashed"):tr("Solid")]])+`<div class="sectionTitle">${esc(tr("Frames"))}</div>${frameList(fs.slice(0,30))}`;renderList();draw();}
function renderStats(){stats.innerHTML=[[tr("Nodes"),summary.node_count||graph.nodes.length],[tr("Edges"),summary.edge_count||graph.edges.length],[tr("Frames"),summary.frame_count||graph.frames.length]].map(([k,v])=>`<div class="stat"><b>${esc(v)}</b><span>${esc(k)}</span></div>`).join("");sourcePath.textContent="zhouzhichao@hrbeu.edu.cn";stageSummary.textContent=`${graph.nodes.length} ${tr("devices")}, ${graph.edges.length} ${tr("links")}`;badges.innerHTML=[`${tr("Output")}: ${graph.source_xlsx||"output.xlsx"}`,`${tr("Capture files")}: ${(summary.capture_files||[]).length||1}`,`${tr("3D view")}: index_3D.html`].map(x=>`<span class="badge">${esc(x)}</span>`).join("");}
function renderLegend(){legend.innerHTML=["Access Point","Hidden AP","Station","Broadcast"].map(k=>`<span class="chip"><span class="dot" style="background:${colorKind(k)}"></span>${esc(kindText(k))}</span>`).join("");}
function hasDisplaySsid(n){const ssid=displaySsid(n.ssid,"");return !!(ssid&&ssid!=="Hidden WiFi");}
function listPrimaryName(n){if(n.kind==="Broadcast")return kindText("Broadcast");if(isApNode(n))return hasDisplaySsid(n)?displaySsid(n.ssid,""):(n.label||n.id);return n.id||n.label||"";}
function listMeta(n){if(n.kind==="Station")return`${kindText(n.kind)} · ${displaySsid(n.ssid,tr("no SSID"))}`;return kindText(n.kind);}
function listRank(n){if(n.kind==="Broadcast")return 0;if(isApNode(n)&&hasDisplaySsid(n))return 1;if(isApNode(n))return 2;return 3;}
function listItemClass(n){return["nodeItem",selectedType==="node"&&selectedId===n.id?"active":"",nodeAddressRole(n)].filter(Boolean).join(" ");}
function renderListItem(n){return`<button class="${listItemClass(n)}" style="--node-color:${colorKind(n.kind)}" data-id="${esc(n.id)}"><span class="swatch" style="background:${colorKind(n.kind)}"></span><span><span class="name">${esc(listPrimaryName(n))}</span><span class="meta">${esc(listMeta(n))}</span></span><span class="count">${esc(n.frames||0)}</span></button>`;}
function renderNodeColumn(title,items){return`<div class="nodeColumn"><div class="nodeColumnHead"><span>${esc(tr(title))}</span><b>${esc(items.length)}</b></div>${items.map(renderListItem).join("")||`<div class="empty">${esc(tr("No nodes."))}</div>`}</div>`;}
function renderList(){const q=filter.value.trim().toLowerCase(),matches=n=>!q||[n.id,n.label,n.kind,n.vendor,n.ssid,n.bssid].join(" ").toLowerCase().includes(q),sorter=(a,b)=>listRank(a)-listRank(b)||kindRank(a)-kindRank(b)||String(listPrimaryName(a)).localeCompare(String(listPrimaryName(b)))||Number(b.frames||0)-Number(a.frames||0)||String(a.id).localeCompare(String(b.id));const ap=graph.nodes.filter(n=>matches(n)&&(n.kind==="Broadcast"||isApNode(n))).sort(sorter),station=graph.nodes.filter(n=>matches(n)&&n.kind==="Station").sort(sorter),other=graph.nodes.filter(n=>matches(n)&&n.kind!=="Broadcast"&&!isApNode(n)&&n.kind!=="Station").sort(sorter);nodeList.innerHTML=renderNodeColumn("AP",ap.concat(other))+renderNodeColumn("Station",station);nodeList.querySelectorAll(".nodeItem").forEach(el=>{el.addEventListener("click",()=>handleListClick(el.dataset.id));});}
function closeLocalTopology(){const had=localTopology.classList.contains("show")||selectedType;if(!had)return false;selectedType="";selectedId="";localTopology.classList.remove("show");localTopologySvg.innerHTML="";localTopologyMeta.textContent="";renderList();draw();return true;}
function overview(){selectedType="";selectedId="";localTopology.classList.remove("show");setDetailHeader(tr("Topology Overview"),"Overview");details.className="panel";const top=[...graph.nodes].sort((a,b)=>Number(b.frames||0)-Number(a.frames||0)).slice(0,8);details.innerHTML=metrics([[tr("AP / Hidden AP"),graph.nodes.filter(n=>isApNode(n)).length],[tr("Intel Stations"),graph.nodes.filter(n=>n.vendor==="Intel").length],[tr("Broadcast"),graph.nodes.filter(n=>n.kind==="Broadcast").length]])+`<div class="sectionTitle">${esc(tr("Most Active Nodes"))}</div>`+top.map(n=>`<div class="frame"><b>${esc(n.label||n.id)}</b><br>${esc(n.vendor?`${n.vendor} ${kindText(n.kind)}`:kindText(n.kind))} / ${esc(n.frames||0)} ${esc(tr("Frames"))} / ${esc(tr("Degree"))} ${degree(n.id)}</div>`).join("");}
function setPlot(img,n,which){const fig=img.closest(".plotFigure");fig.classList.remove("missing");img.onload=()=>fig.classList.remove("missing");img.onerror=()=>fig.classList.add("missing");img.alt=`${n.id} ${which} plot`;img.src=plotPath(n,which);}
function openPlots(id){const n=nodeMap.get(id);if(!n)return;document.getElementById("modalTitle").textContent=n.label||n.id;document.getElementById("modalSub").textContent=`${kindText(n.kind)} / ${n.frames||0} ${tr("Frames")}`;setPlot(document.getElementById("timelinePlot"),n,"timeline");setPlot(document.getElementById("framePlot"),n,"frame");document.getElementById("plotModal").classList.add("show");document.getElementById("plotModal").setAttribute("aria-hidden","false");}
function closePlots(){document.getElementById("plotModal").classList.remove("show");document.getElementById("plotModal").setAttribute("aria-hidden","true");}
function renderAbbrTable(){document.getElementById("abbrTable").innerHTML=`<table><thead><tr><th>${esc(tr("Abbr."))}</th><th>${esc(tr("Full Name"))}</th><th>${esc(tr("Chinese Translation"))}</th></tr></thead><tbody>${abbrItems.map(([a,b,c])=>`<tr><td>${esc(a)}</td><td>${esc(b)}</td><td>${esc(c)}</td></tr>`).join("")}</tbody></table>`;}
function renderChannelWaveforms(){const grid=document.getElementById("channelWaveformGrid"),items=graph.channel_plots||[];grid.innerHTML=items.length?items.map((item,i)=>`<figure class="channelCard"><figcaption><b>${esc(item.label||item.file||`Channel ${i+1}`)}</b><span>${esc(item.sample_rate_msps||20)} Msps / ${esc(item.duration_s||1)} s</span></figcaption><img src="${esc(item.path)}" alt="${esc(item.label||item.file||"channel waveform")}"></figure>`).join(""):`<p class="channelEmpty">${esc(tr("No channel waveform images found. Regenerate topology with plot generation enabled."))}</p>`;}
function renderFrameLegend(){document.getElementById("frameLegend").innerHTML=`<h3>${esc(tr("Node Frame Meaning"))}</h3><div class="frameLegendGrid"><div class="frameLegendItem"><div class="frameLegendSwatch sourceOnly"></div><b>${esc(tr("Source only"))}</b><p>${esc(tr("Black solid frame: the node appears only as a source address, meaning radiated signal is detected from this device."))}</p></div><div class="frameLegendItem"><div class="frameLegendSwatch sourceTarget"></div><b>${esc(tr("Source and target"))}</b><p>${esc(tr("Gold solid frame: the node appears as both source and target address."))}</p></div><div class="frameLegendItem"><div class="frameLegendSwatch targetOnly"></div><b>${esc(tr("Target only"))}</b><p>${esc(tr("Dashed frame: the node is identified only from target address, without corresponding radiated signal."))}</p></div></div>`;}
function openAbbrModal(){renderAbbrTable();renderChannelWaveforms();renderFrameLegend();document.getElementById("abbrModal").classList.add("show");document.getElementById("abbrModal").setAttribute("aria-hidden","false");}
function closeAbbrModal(){document.getElementById("abbrModal").classList.remove("show");document.getElementById("abbrModal").setAttribute("aria-hidden","true");}
function refreshStaticText(){filter.placeholder=tr("Search MAC / SSID / type");document.querySelector(".stageTitle b").textContent=tr("3D Decoded Device Graph");document.querySelector(".localTop b").textContent=tr("Local Topology");document.querySelector(".hint").textContent=tr("Drag empty space to rotate. Arrow keys or WASD move the camera viewpoint. Shift moves faster. Drag a node to move it.");document.getElementById("zoomOut").title=tr("Zoom out");document.getElementById("zoomIn").title=tr("Zoom in");document.getElementById("zoomReset").title=tr("Reset 3D view");document.getElementById("zoomReset").textContent=tr("Reset");document.getElementById("closeModal").textContent=tr("Close");document.getElementById("closeAbbrModal").textContent=tr("Close");document.querySelector("#plotModal .plotFigure:first-child figcaption").textContent=tr("Extracted Signal Timeline");document.querySelector("#plotModal .plotFigure:nth-child(2) figcaption").textContent=tr("Single Frame Detail");document.querySelector("#plotModal .plotFigure:first-child .missingNote").textContent=tr("No timeline image found in signal_plots.");document.querySelector("#plotModal .plotFigure:nth-child(2) .missingNote").textContent=tr("No frame image found in signal_plots.");document.querySelector("#abbrModal .modalTop b").textContent=tr("Abbreviations");document.querySelector("#abbrModal .modalTop span").textContent=tr("English full name and Chinese translation");document.querySelectorAll(".langOption").forEach(btn=>btn.classList.toggle("active",btn.dataset.lang===uiLang));if(document.getElementById("abbrModal").classList.contains("show")){renderAbbrTable();renderChannelWaveforms();renderFrameLegend();}}
function setLanguage(lang){uiLang=lang==="en"?"en":"zh";refreshStaticText();renderStats();renderLegend();renderList();if(localTopology.classList.contains("show")&&selectedType==="node"&&nodeMap.has(selectedId))renderLocalTopology(selectedId);if(selectedType==="node"&&nodeMap.has(selectedId))selectNode(selectedId);else if(selectedType==="edge"&&edgeMap.has(selectedId))selectEdge(selectedId);else overview();draw();}
function isQuickRepeat(last,id){return last.id===id&&Date.now()-last.time<460;}
function handleNodeTap(id){if(isQuickRepeat(lastNodeClick,id)){clearTimeout(pendingNodeClickTimer);pendingNodeClickTimer=0;lastNodeClick={id:"",time:0};openPlots(id);return;}clearTimeout(pendingNodeClickTimer);lastNodeClick={id,time:Date.now()};pendingNodeClickTimer=setTimeout(()=>{if(lastNodeClick.id===id){lastNodeClick={id:"",time:0};selectNode(id);}pendingNodeClickTimer=0;},460);}
function handleListClick(id){if(isQuickRepeat(lastListClick,id)){clearTimeout(pendingListClickTimer);pendingListClickTimer=0;lastListClick={id:"",time:0};openPlots(id);return;}clearTimeout(pendingListClickTimer);lastListClick={id,time:Date.now()};pendingListClickTimer=setTimeout(()=>{if(lastListClick.id===id){lastListClick={id:"",time:0};selectNode(id);}pendingListClickTimer=0;},460);}
function startNodeDrag(evt,id,p){if(evt.button!==0)return;evt.stopPropagation();nodeDrag={id,pointerId:evt.pointerId,startClientX:evt.clientX,startClientY:evt.clientY,start:{...p},moved:false};svg.setPointerCapture(evt.pointerId);svg.classList.add("panning");}
function finishNodeDrag(evt,id){if(suppressNodeClick===id){suppressNodeClick="";evt.preventDefault();evt.stopPropagation();return false;}if(nodeDrag&&nodeDrag.id===id&&nodeDrag.moved){evt.preventDefault();evt.stopPropagation();return false;}return true;}
svg.addEventListener("pointerdown",evt=>{if(evt.button!==0)return;if(evt.target.closest(".node,.apNode,.iconNode,.edge,.edgeHit"))return;dragState={id:evt.pointerId,x:evt.clientX,y:evt.clientY,angleX:view.angleX,angleY:view.angleY};svg.setPointerCapture(evt.pointerId);svg.classList.add("panning");});
svg.addEventListener("pointermove",evt=>{if(nodeDrag&&evt.pointerId===nodeDrag.pointerId){const dx=evt.clientX-nodeDrag.startClientX,dy=evt.clientY-nodeDrag.startClientY;if(Math.abs(dx)>3||Math.abs(dy)>3)nodeDrag.moved=true;const b=screenBasis(),amount=1.15/Math.max(.35,view.zoom);manualWorld.set(nodeDrag.id,{x:nodeDrag.start.x+(b.right.x*dx+b.up.x*dy)*amount,y:nodeDrag.start.y+(b.right.y*dx+b.up.y*dy)*amount,z:nodeDrag.start.z+(b.right.z*dx+b.up.z*dy)*amount});draw();return;}if(!dragState||evt.pointerId!==dragState.id)return;view.angleY=dragState.angleY+(evt.clientX-dragState.x)*.006;view.angleX=clamp(dragState.angleX+(evt.clientY-dragState.y)*.006,-1.35,1.18);draw();});
svg.addEventListener("pointerup",evt=>{if(nodeDrag&&evt.pointerId===nodeDrag.pointerId){const id=nodeDrag.id,moved=nodeDrag.moved;nodeDrag=null;svg.classList.remove("panning");evt.preventDefault();evt.stopPropagation();if(!moved)handleNodeTap(id);return;}if(dragState&&evt.pointerId===dragState.id){dragState=null;svg.classList.remove("panning");}});
svg.addEventListener("pointercancel",()=>{dragState=null;nodeDrag=null;svg.classList.remove("panning");});
svg.addEventListener("wheel",evt=>{evt.preventDefault();view.zoom=clamp(view.zoom*(evt.deltaY<0?1.12:.89),.35,3.8);draw();},{passive:false});
document.getElementById("zoomIn").addEventListener("click",()=>{view.zoom=clamp(view.zoom*1.22,.35,3.8);draw();});
document.getElementById("zoomOut").addEventListener("click",()=>{view.zoom=clamp(view.zoom*.82,.35,3.8);draw();});
document.getElementById("zoomReset").addEventListener("click",()=>{view={angleX:.46,angleY:.72,zoom:1,cameraX:0,cameraZ:0};manualWorld.clear();draw();});
document.getElementById("closeModal").addEventListener("click",closePlots);
document.getElementById("plotModal").addEventListener("click",evt=>{if(evt.target.id==="plotModal")closePlots();});
document.getElementById("openAbbrModal").addEventListener("click",openAbbrModal);
document.getElementById("closeAbbrModal").addEventListener("click",closeAbbrModal);
document.getElementById("abbrModal").addEventListener("click",evt=>{if(evt.target.id==="abbrModal")closeAbbrModal();});
document.querySelectorAll(".langOption").forEach(btn=>btn.addEventListener("click",()=>setLanguage(btn.dataset.lang)));
function moveCameraOnGround(key,fast){const step=(fast?120:48)/Math.max(.55,view.zoom),b=horizontalBasis();if(key==="ArrowLeft"){view.cameraX-=b.right.x*step;view.cameraZ-=b.right.z*step;}else if(key==="ArrowRight"){view.cameraX+=b.right.x*step;view.cameraZ+=b.right.z*step;}else if(key==="ArrowUp"){view.cameraX+=b.forward.x*step;view.cameraZ+=b.forward.z*step;}else if(key==="ArrowDown"){view.cameraX-=b.forward.x*step;view.cameraZ-=b.forward.z*step;}draw();}
function cameraMoveKey(key){const k=String(key||"").toLowerCase();if(k==="w")return"ArrowUp";if(k==="s")return"ArrowDown";if(k==="a")return"ArrowLeft";if(k==="d")return"ArrowRight";return["ArrowUp","ArrowDown","ArrowLeft","ArrowRight"].includes(key)?key:"";}
document.addEventListener("keydown",evt=>{if(evt.key==="Escape"){const plot=document.getElementById("plotModal"),abbr=document.getElementById("abbrModal");if(plot.classList.contains("show")){closePlots();evt.preventDefault();return;}if(abbr.classList.contains("show")){closeAbbrModal();evt.preventDefault();return;}if(closeLocalTopology()){evt.preventDefault();return;}return;}const moveKey=cameraMoveKey(evt.key);if(moveKey&&!["INPUT","TEXTAREA","SELECT"].includes(document.activeElement?.tagName)){evt.preventDefault();moveCameraOnGround(moveKey,evt.shiftKey);}});
filter.addEventListener("input",renderList);window.addEventListener("resize",draw);
setLanguage("zh");
</script>
</body>
</html>"""
    return template.replace("__GRAPH_JSON__", graph_json)

# 将拓扑数据写成 index.html，默认放在 decode_usrp_wifi.py 同一个目录。
def write_topology_html(
    results: List[ResultEntry],
    capture_file: Path,
    topology_dir: Path,
    max_complex: Optional[int] = None,
    generate_plots: bool = True,
) -> Path:
    topology_dir.mkdir(parents=True, exist_ok=True)
    copy_topology_icons(topology_dir)
    plot_paths = generate_node_signal_plots(results, capture_file, topology_dir, max_complex) if generate_plots else {}
    channel_plots = generate_channel_waveform_plots([capture_file], topology_dir, max_complex) if generate_plots else []
    graph = build_topology(results, capture_file, plot_paths, channel_plots)
    output = topology_dir / "index_3D.html"
    output.write_text(topology_html_3d(graph), encoding="utf-8")
    return output

# 多文件批量解析后的拓扑写出：节点和边汇总所有 .bin，节点信号图按第一次出现的文件生成。
def write_topology_html_multi(
    capture_results: Sequence[CaptureResult],
    capture_files: Sequence[Path],
    topology_dir: Path,
    max_complex: Optional[int] = None,
    generate_plots: bool = True,
) -> Path:
    topology_dir.mkdir(parents=True, exist_ok=True)
    copy_topology_icons(topology_dir)
    plot_paths = generate_node_signal_plots_multi(capture_results, topology_dir, max_complex) if generate_plots else {}
    channel_plots = generate_channel_waveform_plots(capture_files, topology_dir, max_complex) if generate_plots else []
    source_dir = capture_files[0].parent if capture_files else topology_dir
    graph = build_topology_multi(capture_results, capture_files, plot_paths, source_dir, channel_plots)
    output = topology_dir / "index_3D.html"
    output.write_text(topology_html_3d(graph), encoding="utf-8")
    return output

def column_name(idx: int) -> str:
    name = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        name = chr(65 + rem) + name
    return name

def column_index(name: str) -> int:
    col = 0
    for ch in name:
        col = col * 26 + ord(ch) - 64
    return col - 1

# 只用 Python 标准库手工写入 .xlsx 文件，避免依赖 pandas/openpyxl。
def write_xlsx(path: Path, rows: List[List[object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    sheet_rows = [HEADERS] + rows
    shared: Dict[str, int] = {}
    shared_values: List[str] = []

    # 给共享字符串表中的文本分配编号，减少 xlsx 内部重复字符串。
    def shared_index(text: str) -> int:
        if text not in shared:
            shared[text] = len(shared_values)
            shared_values.append(text)
        return shared[text]

    row_xml = []
    for r_idx, row in enumerate(sheet_rows, 1):
        cells = []
        for c_idx, value in enumerate(row):
            ref = f"{column_name(c_idx)}{r_idx}"
            if value is None or value == "":
                continue
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                cells.append(f'<c r="{ref}"><v>{value}</v></c>')
            else:
                idx = shared_index(str(value))
                cells.append(f'<c r="{ref}" t="s"><v>{idx}</v></c>')
        row_xml.append(f'<row r="{r_idx}">{"".join(cells)}</row>')

    xml_head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    shared_xml = "".join(f'<si><t>{escape(text)}</t></si>' for text in shared_values)
    files = {
        "[Content_Types].xml": (
            f'{xml_head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
        ),
        "_rels/.rels": (
            f'{xml_head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>'
        ),
        "xl/workbook.xml": (
            f'{xml_head}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            f'{xml_head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
        ),
        "xl/worksheets/sheet1.xml": (
            f'{xml_head}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<sheetData>{"".join(row_xml)}</sheetData></worksheet>'
        ),
        "xl/sharedStrings.xml": (
            f'{xml_head}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'count="{len(shared_values)}" uniqueCount="{len(shared_values)}">{shared_xml}</sst>'
        ),
        "xl/styles.xml": (
            f'{xml_head}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs></styleSheet>'
        ),
    }
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)

# 从已有 xlsx 提示表中读取起始采样点、MCS 和 PSDU 长度，用于下一次加速解码。
def read_hint_entries(path: Path) -> List[HintEntry]:
    if not path.exists() or path.suffix.lower() != ".xlsx":
        return []
    try:
        with zipfile.ZipFile(path) as z:
            shared: List[str] = []
            if "xl/sharedStrings.xml" in z.namelist():
                text = z.read("xl/sharedStrings.xml").decode("utf-8", errors="replace")
                shared = ["".join(re.findall(r"<t[^>]*>(.*?)</t>", si, flags=re.S))
                          for si in re.findall(r"<si>(.*?)</si>", text, flags=re.S)]
            sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8", errors="replace")

        rows: List[Dict[int, str]] = []
        for row_xml in re.findall(r"<row[^>]*>(.*?)</row>", sheet, flags=re.S):
            cells: Dict[int, str] = {}
            for attrs, body in re.findall(r"<c([^>]*)>(.*?)</c>", row_xml, flags=re.S):
                ref = re.search(r'r="([A-Z]+)(\d+)"', attrs)
                val = re.search(r"<v>(.*?)</v>", body)
                if ref and val:
                    raw = val.group(1)
                    cells[column_index(ref.group(1))] = shared[int(raw)] if 't="s"' in attrs else raw
            rows.append(cells)

        if not rows:
            return []
        header = {v: k for k, v in rows[0].items()}
        start_col = header.get("起始采样点")
        if start_col is None:
            return []

        mcs_col = header.get("MCS")
        len_col = header.get("PSDU长度")
        dedup: Dict[int, HintEntry] = {}
        for row in rows[1:]:
            start = row.get(start_col, "")
            if not str(start).isdigit():
                continue
            mcs = row.get(mcs_col, "") if mcs_col is not None else ""
            length = row.get(len_col, "") if len_col is not None else ""
            dedup[int(start)] = HintEntry(
                int(start),
                int(mcs) if str(mcs).isdigit() else None,
                int(length) if str(length).isdigit() else None,
            )
        return sorted(dedup.values(), key=lambda item: item.start_sample)
    except Exception:
        return []

def count_unique_ssids(results: List[ResultEntry]) -> int:
    return len({r.ssid for r in results if r.ssid})

def console_text(text: object) -> str:
    return str(text).encode("gbk", errors="backslashreplace").decode("gbk")

# 解析命令行参数，默认输出固定为脚本目录下的 output.xlsx。
def parse_args() -> argparse.Namespace:
    script_dir = Path(__file__).resolve().parent
    default_capture_dir = script_dir / "adc signal"
    if not default_capture_dir.exists() and (script_dir.parent / "adc signal").exists():
        default_capture_dir = script_dir.parent / "adc signal"
    parser = argparse.ArgumentParser(description="USRP B210 Wi-Fi MAC node decode plus, Python version")
    parser.add_argument("--capture", default=os.environ.get("WIFI_CAPTURE_FILE", ""), help="Optional single USRP sc16 .bin capture file")
    parser.add_argument("--capture-dir", default=os.environ.get("WIFI_CAPTURE_DIR", str(default_capture_dir)), help="Directory containing .bin capture files")
    parser.add_argument("--output", default=os.environ.get("WIFI_OUTPUT_FILE", str(script_dir / "output.xlsx")), help="Output .xlsx file")
    parser.add_argument("--topology-dir", default=os.environ.get("WIFI_TOPOLOGY_DIR", str(script_dir)), help="Directory for generated index_3D.html topology view")
    parser.add_argument("--no-topology", action="store_true", help="Decode only; do not generate topology index_3D.html")
    parser.add_argument("--no-plots", action="store_true", help="Skip per-node signal plot PNG generation for faster batch runs")
    parser.add_argument("--max-complex", type=int, default=None, help="Maximum complex samples to read")
    parser.add_argument("--max-packets", type=int, default=None, help="Maximum packet candidates to inspect")
    parser.add_argument("--max-frames-per-file", type=int, default=int(env_number("WIFI_MAX_FRAMES_PER_FILE", 0) or 0), help="Stop decoding a file after this many frames; 0 means no limit")
    parser.add_argument("--max-psdu-length", type=int, default=int(env_number("WIFI_MAX_PSDU_LENGTH", 800) or 800), help="Skip implausibly long PSDU candidates")
    parser.add_argument("--energy-gate-factor", type=float, default=float(env_number("WIFI_ENERGY_GATE_FACTOR", 3.0) or 3.0), help="Skip quiet windows below this relative energy factor; <=0 disables")
    parser.add_argument("--inactive-file-factor", type=float, default=float(env_number("WIFI_INACTIVE_FILE_FACTOR", 8.0) or 8.0), help="Skip whole files below this activity factor; <=0 disables")
    parser.add_argument("--no-skip-inactive", action="store_true", help="Do not skip low-activity capture files")
    parser.add_argument("--lltf-max-candidates", type=int, default=int(env_number("WIFI_LLTF_MAX_CANDIDATES", 64) or 64), help="Maximum L-LTF candidates per detection window")
    parser.add_argument("--stf-max-candidates", type=int, default=int(env_number("WIFI_STF_MAX_CANDIDATES", 64) or 64), help="Maximum STF candidates per detection window")
    parser.add_argument("--hints", default="", help="Optional XLSX file with 起始采样点 column for faster decoding")
    parser.add_argument("--use-hints", action="store_true", help="Use automatic hint loading from capture/output .xlsx files")
    parser.add_argument("--no-hints", action="store_true", help="Disable automatic hint loading from output.xlsx")
    parser.add_argument("--fallback-on-miss", action="store_true", help="Try additional length/MCS candidates when fast decode misses")
    parser.add_argument("--no-control-rescue", action="store_true", help="Disable lightweight RTS/CTS/BlockAck control-frame rescue attempts")
    parser.add_argument("--try-both-byte-orders", action="store_true", help="Try both PSDU bit-to-byte orders")
    parser.add_argument("--exhaustive-mcs", action="store_true", help="Try all legacy OFDM MCS values")
    parser.add_argument("--scan-step-samples", type=int, default=int(env_number("WIFI_SCAN_STEP_SAMPLES", 320) or 320), help="Samples to advance after each inspected packet candidate")
    parser.add_argument("--verbose-frames", action="store_true", help="Print decoded frames as they are found")
    return parser.parse_args()

# 主流程入口：读取参数、加载提示、执行解码，并且只写出 output.xlsx。
def main() -> int:
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    capture_files = resolve_captures(script_dir, args.capture, args.capture_dir)
    output_file = Path(args.output)

    opt = DecodeOptions(
        fast_mode=env_flag("WIFI_FAST_MODE", True),
        try_both_byte_orders=args.try_both_byte_orders or env_flag("WIFI_TRY_BOTH_BYTE_ORDERS", False),
        exhaustive_mcs=args.exhaustive_mcs or env_flag("WIFI_EXHAUSTIVE_MCS", False),
        verbose_frames=args.verbose_frames or env_flag("WIFI_VERBOSE_FRAMES", False),
        progress_step_percent=float(env_number("WIFI_PROGRESS_STEP_PERCENT", 5) or 5),
        detect_window_samples=int(env_number("WIFI_DETECT_WINDOW_SAMPLES", 40_000) or 40_000),
        detect_overlap_samples=int(env_number("WIFI_DETECT_OVERLAP_SAMPLES", 4_000) or 4_000),
        energy_gate_factor=args.energy_gate_factor,
        lltf_max_candidates=max(1, args.lltf_max_candidates),
        stf_max_candidates=max(1, args.stf_max_candidates),
        fallback_on_miss=args.fallback_on_miss or env_flag("WIFI_FALLBACK_ON_MISS", False),
        max_complex_samples=args.max_complex or (int(env_number("WIFI_MAX_COMPLEX", 0) or 0) or None),
        max_packets=args.max_packets or (int(env_number("WIFI_MAX_PACKETS", 0) or 0) or None),
        max_frames=args.max_frames_per_file or None,
        max_psdu_length=args.max_psdu_length,
        scan_step_samples=max(80, args.scan_step_samples),
        rescue_control_frames=not args.no_control_rescue,
    )

    print("USRP B210 WiFi MAC node decode plus (Python)")
    print(f"capture_count: {len(capture_files)}")
    print(f"capture_dir: {Path(args.capture_dir) if not args.capture else capture_files[0].parent}")
    print("format: sc16")
    print(f"sample_rate: {SAMPLE_RATE:.0f} Hz\n")
    start_time = time.time()
    capture_results: List[CaptureResult] = []
    empty_captures: List[Path] = []
    for file_idx, capture_file in enumerate(capture_files, 1):
        print(f"\n[{file_idx}/{len(capture_files)}] capture: {capture_file}")
        if not args.no_skip_inactive and not capture_has_activity(capture_file, args.inactive_file_factor):
            print(f"skip_inactive_capture: {capture_file.name}")
            empty_captures.append(capture_file)
            continue
        hint_entries: List[HintEntry] = []
        # 默认全信号扫描。旧 output.xlsx 只能提示已经解析过的起点，如果旧结果漏帧，
        # 自动复用提示会让漏帧固化；需要快速复现旧起点时再显式打开 --use-hints。
        use_hints = (args.use_hints or bool(args.hints) or env_flag("WIFI_USE_HINTS", False)) and not args.no_hints
        if use_hints:
            hint_files = [Path(args.hints)] if args.hints else [
                capture_file.with_suffix(".xlsx"),
                capture_file.with_name("output.xlsx"),
            ]
            for hint_file in hint_files:
                hint_entries = read_hint_entries(hint_file)
                if hint_entries:
                    print(f"using_hint_file: {hint_file}")
                    break
        results = decode_known_starts(capture_file, opt, hint_entries) if hint_entries else decode_capture(capture_file, opt)
        if args.max_frames_per_file and len(results) > args.max_frames_per_file:
            results = results[:args.max_frames_per_file]
        if results:
            capture_results.extend(CaptureResult(capture_file, result) for result in results)
        else:
            empty_captures.append(capture_file)
    rows = result_rows_multi(capture_results, empty_captures)
    write_xlsx(output_file, rows)
    topology_file = None if args.no_topology else write_topology_html_multi(
        capture_results, capture_files, Path(args.topology_dir), opt.max_complex_samples, generate_plots=not args.no_plots
    )
    elapsed = time.time() - start_time

    all_results = [item.result for item in capture_results]
    print(f"\nSUMMARY files={len(capture_files)} frames={len(all_results)} fcs_ok={sum(r.fcs_ok for r in all_results)} unique_ssids={count_unique_ssids(all_results)} empty_files={len(empty_captures)}")
    if not all_results:
        print("No reliable WiFi MAC frame was decoded. Try another channel capture or a longer capture.")
    else:
        for k, item in enumerate(capture_results[:20], 1):
            r = item.result
            print(console_text(
                f"{k:2d}. type={r.frame_type}/{r.frame_subtype}  source={r.source_node}  "
                f"destination={r.destination_node}  bssid={r.bssid}  ssid={r.ssid}  "
                f"start={r.start_sample}  fcs={r.fcs_ok}  file={item.capture_file.name}"
            ))
        if len(capture_results) > 20:
            print(f"... {len(capture_results) - 20} more frames were saved to Excel.")
    print(f"\nExcel output was saved to: {output_file}")
    if topology_file:
        print(f"Topology HTML was saved to: {topology_file}")
    print(f"elapsed_seconds={elapsed:.2f}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
