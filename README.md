# usrp-wifi-topology-decoder

Python tools for decoding WiFi MAC frames from USRP B210 `sc16` IQ captures and visualizing the inferred device topology.

The project reads one or more raw WiFi channel captures, detects legacy 802.11 OFDM packets, extracts MAC-layer information, writes the decoded frame table to `output.xlsx`, and generates an interactive topology view as HTML.

<img width="1912" height="922" alt="6437ec9b4847a3e404e08c299340fc08" src="https://github.com/user-attachments/assets/9f8ba45c-f209-4351-bb35-de8c6487444f" />

<img width="1912" height="922" alt="d4a1e6353104f29a6f12545f127ee756" src="https://github.com/user-attachments/assets/9135baae-253e-4306-aae0-b1ac8e2770f1" />

<img width="1912" height="922" alt="611e650d778c4fb7cdcde42dcdd1b4b3" src="https://github.com/user-attachments/assets/521418d2-cc36-432d-929d-60668f90b8eb" />


## Features

- Decode USRP B210 complex `sc16` IQ recordings.
- Support batch processing of all `.bin` files in an `adc signal/` folder.
- Extract SSID, BSSID, source address, destination address, receiver address, transmitter address, frame type, frame subtype, sequence number, MCS, PSDU length, and FCS status.
- Build a device graph from decoded source and target relationships.
- Generate per-node signal plots:
  - a large-scale extracted timeline over the full capture duration,
  - a single-frame waveform detail view.
- Generate 2D or 3D interactive topology HTML pages.
- Use built-in Intel OUI prefixes for offline station vendor identification.
- Use local router and laptop icons from `assets/` when available.


  <img width="1595" height="435" alt="01_wifi_2g4_ch001_2412mhz_20msps_1s_rx2_sc16_waveform" src="https://github.com/user-attachments/assets/9deb445e-a44b-459b-8fd5-132c35a7952d" />

<img width="1575" height="480" alt="cc_2d_21_e5_c1_a4_frame" src="https://github.com/user-attachments/assets/7d30cdf4-cf72-4871-8e94-cd3b3fac745a" />

<img width="1575" height="480" alt="cc_2d_21_e5_c1_a4_timeline" src="https://github.com/user-attachments/assets/6d8d5d2c-a4dc-4db0-8f27-bcd68ca52a1f" />


## Repository Layout

The minimal release package contains the Python decoder and an ADC-signal folder:

```text
.
|-- decode_usrp_wifi_3D.py       # Python decoder and 3D topology visualizer
`-- adc signal/                  # Example USRP WiFi IQ captures
```

The Python program automatically parses the `.bin` signals under the ADC-signal folder. Because of file-transfer limits, only a few example captures are included in this repository. The complete capture set can be downloaded from Baidu Netdisk:

```text
https://pan.baidu.com/s/1mOyiQMpmzp66i20ioPoepQ?pwd=9921
```

After a successful run, the program outputs an Excel file and an HTML file. Open the generated HTML file in a browser to view the interactive WiFi topology visualization interface.

A full working directory may look like this after running the decoder:

```text
.
|-- decode_usrp_wifi.py          # Main 2D decoder and topology generator
|-- decode_usrp_wifi_3D.py       # 3D topology version
|-- assets/                      # Optional node icons used by the HTML views
|-- adc signal/                  # Put input .bin capture files here
|-- signal_plots/                # Generated node and channel waveform images
|-- output.xlsx                  # Generated decode result table
|-- index.html                   # Generated 2D topology page
`-- index_3D.html                # Generated 3D topology page
```

Generated files such as `output.xlsx`, `index.html`, `index_3D.html`, `signal_plots/`, and large `.bin` captures do not need to be committed unless you intentionally want to publish examples.

## Requirements

- Python 3.9 or newer
- NumPy
- Matplotlib, optional but recommended for signal plot generation

Install the Python dependencies:

```bash
pip install numpy matplotlib
```

The scripts write `.xlsx` files using Python standard-library ZIP/XML utilities, so `openpyxl` is not required.

## Input Data

The decoder expects raw USRP `sc16` IQ data:

- interleaved signed 16-bit integers,
- I/Q order: `I0, Q0, I1, Q1, ...`,
- sample rate: 20 Msps,
- intended for 20 MHz legacy 802.11a/g OFDM packets.

By default, place capture files in:

```text
adc signal/
```

Each file should use the `.bin` extension. Temporary files whose names start with `debug_`, `test_`, or `seq` are skipped by default.

## Quick Start

Run the 2D decoder:

```bash
python decode_usrp_wifi.py
```

This produces:

```text
output.xlsx
index.html
signal_plots/
```

Run the 3D decoder:

```bash
python decode_usrp_wifi_3D.py
```

This produces:

```text
output.xlsx
index_3D.html
signal_plots/
```

Open the generated HTML file in a browser to inspect the decoded topology.

## Common Commands

Decode one specific capture:

```bash
python decode_usrp_wifi.py --capture path/to/capture.bin
```

Decode all captures in a specific folder:

```bash
python decode_usrp_wifi.py --capture-dir "path/to/adc signal"
```

Generate only `output.xlsx` without topology HTML:

```bash
python decode_usrp_wifi.py --no-topology
```

Skip signal plot generation for faster processing:

```bash
python decode_usrp_wifi.py --no-plots
```

Limit how much data is loaded from each capture:

```bash
python decode_usrp_wifi.py --max-complex 2000000
```

Generate the 3D topology view:

```bash
python decode_usrp_wifi_3D.py
```

## Main Options

| Option | Description |
| --- | --- |
| `--capture` | Decode one specific `.bin` capture file. |
| `--capture-dir` | Decode all `.bin` files in a folder. |
| `--output` | Output Excel file. Default: `output.xlsx`. |
| `--topology-dir` | Directory for generated HTML and plot assets. |
| `--no-topology` | Decode only, without generating topology HTML. |
| `--no-plots` | Skip per-node signal plot PNG generation. |
| `--max-complex` | Maximum number of complex samples to read per capture. |
| `--max-packets` | Maximum number of packet candidates to inspect. |
| `--max-frames-per-file` | Stop decoding a file after this many decoded frames. |
| `--use-hints` | Reuse known packet start hints from existing Excel output. |
| `--hints` | Use a specific XLSX hint file. |
| `--fallback-on-miss` | Try additional length/MCS candidates when fast decoding misses. |
| `--try-both-byte-orders` | Try both PSDU bit-to-byte orders. |
| `--exhaustive-mcs` | Try all legacy OFDM MCS values. |
| `--verbose-frames` | Print decoded frames as they are found. |

The same behavior can also be controlled by environment variables such as `WIFI_CAPTURE_FILE`, `WIFI_CAPTURE_DIR`, `WIFI_OUTPUT_FILE`, `WIFI_TOPOLOGY_DIR`, `WIFI_FAST_MODE`, and `WIFI_MAX_COMPLEX`.

## Output Files

`output.xlsx` contains the decoded frame table, including MAC addresses, SSID/BSSID information, frame type information, timing/sample positions, PHY metadata, and FCS status.

`index.html` provides the 2D topology visualization.

`index_3D.html` provides the 3D topology visualization. The 3D page supports:

- mouse drag to rotate the view,
- mouse wheel or zoom buttons to zoom,
- arrow keys or `W/A/S/D` to move the camera viewpoint,
- dragging nodes to adjust positions,
- double-clicking nodes to open their signal plots,
- selecting nodes to show local topology,
- `Esc` to close the local topology focus,
- Chinese/English UI switching,
- an abbreviation panel with English full names and Chinese translations.

`signal_plots/` contains generated node waveform images and channel-level waveform summaries.

## Topology Semantics

Nodes represent decoded WiFi devices or inferred address entities.

- AP nodes prefer SSID labels when available.
- Station nodes prefer MAC-address labels.
- A station can show the SSID of the AP/network it is associated with.
- Broadcast is treated as an inferred broadcast target, not a physical transmitter.
- Solid links indicate relationships between two decoded radiation/source nodes.
- Dashed links indicate relationships involving target-only inferred nodes.

Node border styles indicate how the node was discovered:

- black solid border: source-only node,
- dashed border: target-only node,
- gold solid border: node appears as both source and target.

When a node has a device icon, the normal black solid frame is hidden for a cleaner view, while dashed or gold frames are still shown when semantically meaningful.

## Notes and Limitations

This decoder focuses on 20 MHz legacy 802.11 OFDM decoding from raw IQ captures. Wideband 40/80 MHz 802.11n/ac/ax captures may still contain decodable legacy or management frames, but full HT/VHT/HE PHY decoding is outside the current scope.

Hidden SSIDs may appear when an AP transmits frames without the SSID information element or when the relevant beacon/probe response is not captured.

FCS failures are useful diagnostics. They can indicate weak signal, collision, timing/frequency offset, unsupported PHY mode, or an incomplete/incorrect packet candidate.

## Citation

If this code or visualization helps your research, please cite the related work:

Z. Zhou, C. Hou, Z. Song, B. Wang, X. Wu and Z. Liu, "Wireless Communication Network Multi-Node Topology Inference Method With Maximizing Information Utilization Based on GCN," in IEEE Transactions on Cognitive Communications and Networking, vol. 12, pp. 4319-4331, 2026.

Zhou Zhichao, Hou Changbo, Meng Guojing, et al. Research on the digital twin of communication relationships in the electromagnetic spectrum of wireless local area networks[J]. Journal on Communications, 2026, 47(05): 78-90.

## Author

Zhou Zhichao  
zhouzhichao@hrbeu.edu.cn
