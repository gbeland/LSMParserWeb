# Samsung MDC Protocol Reference

> [!IMPORTANT]
> This reference is authoritatively derived from `MDC.pdf` (Version 15.4) via direct text extraction.

## Protocol Structure
Responses follow a standard hex structure:
`AA FF [ID] [Length] [Command] [Data...] [Checksum]`

-   **Header**: `AA FF` (Fixed)
-   **ID**: Device ID (e.g., `01` for SBox, `FE` for Broadcast)
-   **Length**: Number of bytes following the Length byte
-   **Command**: The command hex code
-   **Data**: Payload (if any)
-   **Checksum**: Sum of bytes (excluding Header)

## Command List (Comprehensive)

| Command (Hex) | Description | Data Type |
| :--- | :--- | :--- |
| **0x00** | Status Control | Lookup (Power, Volume, Mute, Source, etc.) |
| **0x04** | Image Size Control | Lookup |
| **0x06** | PIP On/Off Control | Boolean |
| **0x08** | Maintenance Control | Variant |
| **0x0A** | Signage Player Control | Variant |
| **0x0B** | Serial Number Control | ASCII |
| **0x0D** | Display Status Control | Lookup (Error Codes) |
| **0x0E** | Software Version Control | ASCII |
| **0x0F** | Auto Motion Plus Control | Lookup |
| **0x10** | Model Number Control | Lookup |
| **0x11** | Power Control | `00`: Off, `01`: On, `02`: Reboot |
| **0x12** | Volume Control | 0-100 |
| **0x13** | Mute Control | `00`: Off, `01`: On |
| **0x14** | Input Source Control | Lookup (HTML, DVI, HDMI, etc.) |
| **0x15** | Picture Size Control | Lookup |
| **0x17** | Display ID Information | Struct |
| **0x18** | Screen Mode Control | Lookup |
| **0x1A** | Outdoor Mode Control | Variant |
| **0x1B** | System Configuration | Variant (Network, Time, Edu) |
| **0x1C** | MagicInfo Control | Variant |
| **0x1D** | MDC Connection Type | Lookup |
| **0x1F** | Still Image Control | Boolean |
| **0x20** | Test Pattern Control | Lookup |
| **0x21** | Picture Control | Sub-commands (Contrast, Brightness, Color) |
| **0x24** | PIP Source Control | Lookup |
| **0x25** | PIP Size Control | Lookup |
| **0x34** | Menu Control | Toggle |
| **0x3D** | Video Wall Mode Control | Boolean |
| **0x3E** | Color Tone Control | Lookup |
| **0x3F** | Color Temperature Control | Decimal (K) |
| **0x4B** | Video Picture Position/Size | Struct |
| **0x50** | Sensor Control | Variant (Light, IR, Temp) |
| **0x56** | Energy Saving Control | Lookup |
| **0x57** | Clock Control | Struct |
| **0x58** | Lamp Control | Manual/Auto |
| **0x5B** | Safety Screen (Burn Protection) | Variant |
| **0x63** | Ticker Control | Variant |
| **0x67** | Device Name Control | ASCII |
| **0x68** | Speaker Select Control | Lookup |
| **0x72** | Sound Mode Control | Lookup |
| **0x84** | Video Wall Mode | Boolean |
| **0x85** | Max Temperature Control | Decimal |
| **0x89** | Video Wall User Control | Variant |
| **0x8A** | Model Name Control | ASCII |
| **0x8C** | Video Wall Feature Control | Variant (Geometry, Layout) |
| **0xB2** | 3/4 Screen Mode Control | Variant |
| **0xB5** | PC Power Control | Boolean |
| **0xC1** | Virtual Remote Key | Lookup (Keys) |
| **0xC2** | User Gamma Control | Struct |
| **0xC4** | Supported Function Info | Bitmask |
| **0xC6** | Eco Solution Control | Variant |
| **0xC8** | OSD Control | Variant |
| **0xC9** | Sound Menu Control | Variant |
| **0xCA** | System Menu Control | Variant |
| **0xD0** | LED Product Feature | Complex Variant |
| **0xFE** | White Balance Control | Complex Variant |

## Data Tables (Select)

### Input Sources (0x14)
| Hex | Source |
| :--- | :--- |
| `14` | PC |
| `18` | DVI |
| `0C` | AV |
| `21` | HDMI1 |
| `23` | HDMI2 |
| `25` | DisplayPort |
| `60` | MagicInfo S |
| `65` | Web Browser |

### Status Codes (0x0D)
| Hex | Error |
| :--- | :--- |
| `00` | Normal |
| `01` | Fan Error |
| `03` | Lamp Error |
| `04` | Brightness Sensor |
| `06` | Source Error |
| `07` | Temperature Error |
| `08` | Panel Error |

### Virtual Remote Keys (0xC1)
| Hex | Key |
| :--- | :--- |
| `1A` | Menu |
| `1B` | Return |
| `60` | Up |
| `61` | Down |
| `62` | Left |
| `65` | Right |
| `68` | Enter |
| `11` | Power |
| `07` | Volume Up |
| `0B` | Volume Down |
