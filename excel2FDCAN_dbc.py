import openpyxl
from pathlib import Path
import sys

# --- Constants ---
EXCEL_FILE = "CAN_FD_DBC_Data.xlsx"
USE_CANFD = True

# --- Helper Functions ---
def hex_to_dec(hex_str):
    """Converts a hex string (like '0xAA') to a decimal integer."""
    try:
        return int(str(hex_str), 16)
    except (ValueError, TypeError):
        return 0

def get_valid_dlc(input_dlc):
    """Finds the smallest valid CAN FD DLC that is >= the input DLC."""
    valid_canfd_lengths = [0, 1, 2, 3, 4, 5, 6, 7, 8, 12, 16, 20, 24, 32, 48, 64]
    if input_dlc is None:
        return 8 # Default DLC
    try:
        dlc = int(input_dlc)
        if dlc in valid_canfd_lengths:
            return dlc
        for valid_dlc in valid_canfd_lengths:
            if valid_dlc >= dlc:
                return valid_dlc
        return 64 # Max length if it exceeds all values
    except (ValueError, TypeError):
        return 8

class CANDBCGenerator:
    """
    Generates a CAN FD DBC file from a specified Excel sheet format.
    """
    def __init__(self, excel_path, use_canfd=True):
        self.excel_path = Path(excel_path)
        self.use_canfd = use_canfd
        self.messages = {}
        self.ecus = set()
        self.ecu_addresses = {}

    def parse_signals(self):
        """Parses messages and signals from the 'Signals' sheet of the Excel file."""
        print(f"모드: {'CAN FD' if self.use_canfd else 'Standard CAN'}")
        print(f"Excel 파일: {self.excel_path}")

        try:
            workbook = openpyxl.load_workbook(self.excel_path)
            sheet = workbook['Signals']
        except (FileNotFoundError, KeyError) as e:
            print(f"오류: '{self.excel_path}' 파일 또는 'Signals' 시트를 찾을 수 없습니다. {e}", file=sys.stderr)
            sys.exit(1)

        print("컬럼 헤더 확인:")
        header = [cell.value for cell in sheet[1]]
        for i, col_name in enumerate(header):
             if i < 28: # Print a reasonable number of headers
                print(f"\t{i}열: '{col_name}'")


        # Read data rows
        for row_index, row_cells in enumerate(sheet.iter_rows(min_row=2)):
            # Convert row_cells to a list of values
            row = [cell.value for cell in row_cells]

            # B열(1)이 비어있으면 건너뜀 (메시지의 시작점으로 간주)
            if len(row) <= 1 or not row[1]:
                continue

            try:
                msg_id_hex = row[6]
                if msg_id_hex is None: continue
                
                # Validate that critical columns are numeric before processing the row
                msg_id = hex_to_dec(msg_id_hex)
                _ = int(row[8]) # Check DLC
                _ = int(row[9]) # Check Cycle Time

                # New message found
                if msg_id not in self.messages:
                    sender_ecu = str(row[1]).strip()
                    source_addr = str(row[2]).strip()
                    
                    self.ecus.add(sender_ecu)
                    if source_addr:
                        self.ecu_addresses[sender_ecu] = source_addr

                    self.messages[msg_id] = {
                        "name": str(row[5]).strip(),
                        "sender": sender_ecu,
                        "node_id": source_addr,
                        "dlc": row[8],
                        "cycle_time": row[9],
                        "cycle_time_fast": row[10],
                        "start_delay_time": row[12],
                        "tx_method": row[25],
                        "frame_type": row[7],
                        "nr_of_repetition": row[73],
                        "il_support": row[74],
                        "nm_message": row[75],
                        "diag_request": row[64],
                        "diag_response": row[65],
                        "signals": []
                    }
                
                # Add signal to the current message if signal name (N열, 13) exists
                if len(row) > 13 and row[13]:
                    # Add receiver ECUs to the global ECU set
                    receivers_raw = row[16]
                    if receivers_raw:
                        receivers = [r.strip() for r in str(receivers_raw).split(',')]
                        self.ecus.update(receivers)
                    self.messages[msg_id]["signals"].append(row)

            except (IndexError, ValueError, TypeError) as e:
                print(f"경고: 행 {row_index + 2} 처리 중 오류 발생, 건너뜁니다. 오류: {e}", file=sys.stderr)
                continue
        
        print("\n메시지:")
        for msg_id, msg_data in self.messages.items():
             print(f"{msg_data['name']} (ID:{hex(msg_id)}), Sender={msg_data['sender']}, Node={msg_data['node_id']}, DLC={msg_data['dlc']}")

        print(f"\n{sum(len(m['signals']) for m in self.messages.values())}개 신호, {len(self.messages)}개 메시지")
        print("ECU Address 설정:")
        for ecu, addr in self.ecu_addresses.items():
            print(f"\t{ecu}: {addr}")


    def generate_dbc(self, output_path):
        """Generates and writes the DBC file content to the output path."""
        if not self.messages:
            print("오류: 생성할 메시지가 없습니다.", file=sys.stderr)
            return

        with open(output_path, 'w', encoding='utf-8') as f:
            # 1. Header
            f.write('VERSION ""\n\n')
            f.write('NS_ :\n\tNS_DESC_\n\tCM_\n\tBA_DEF_\n\tBA_\n\tVAL_\n\tCAT_DEF_\n\tCAT_\n\tFILTER\n\tBA_DEF_DEF_\n\tEV_DATA_\n\tENVVAR_DATA_\n\tSGTYPE_\n\tSGTYPE_VAL_\n\tBA_DEF_SGTYPE_\n\tBA_SGTYPE_\n\tSIG_TYPE_REF_\n\tVAL_TABLE_\n\tSIG_GROUP_\n\tSIG_VALTYPE_\n\tSIGTYPE_VALTYPE_\n\n')
            f.write('BS_:\n\n')

            # 2. ECUs (BU_)
            f.write(f'BU_: {" ".join(sorted(list(self.ecus)))}\n\n')

            # 3. Messages (BO_) and Signals (SG_)
            for msg_id, msg in sorted(self.messages.items()):
                dlc = get_valid_dlc(msg['dlc']) if self.use_canfd else int(msg.get('dlc', 8))
                f.write(f'BO_ {msg_id} {msg["name"]}: {dlc} {msg["sender"]}\n')
                
                for sig_row in msg["signals"]:
                    try:
                        name = sig_row[13]
                        start = int(sig_row[20])
                        length = int(sig_row[18])
                        byte_order = '0' if 'Motorola' in str(sig_row[21]) else '1' # 0=big, 1=little
                        sign = '+' if 'Unsigned' in str(sig_row[40]) else '-'
                        scale = sig_row[44] if sig_row[44] is not None else 1.0
                        offset = sig_row[45] if sig_row[45] is not None else 0.0
                        min_val = sig_row[46] if sig_row[46] is not None else 0
                        max_val = sig_row[47] if sig_row[47] is not None else (2**length) * scale - (1 if sign == '+' else 0)
                        unit = sig_row[49] or ""
                        receivers_raw = sig_row[16]
                        receivers = [r.strip() for r in str(receivers_raw).split(',')] if receivers_raw else ['Vector__XXX']
                        
                        f.write(f' SG_ {name} : {start}|{length}@{byte_order}{sign} ({scale},{offset}) [{min_val}|{max_val}] "{unit}" {",".join(receivers)}\n')
                    except (IndexError, TypeError, ValueError) as e:
                        print(f"경고: 신호 처리 중 오류 (메시지: {msg['name']}), 건너뜁니다. 오류: {e}", file=sys.stderr)
                        continue
                f.write('\n')

            # 4. Comments (CM_)
            for msg_id, msg in self.messages.items():
                for sig_row in msg["signals"]:
                    desc = sig_row[15]
                    if desc:
                        f.write(f'CM_ SG_ {msg_id} {sig_row[13]} "{desc}";\n')
            f.write('\n')

            # 5. Attribute Definitions (BA_DEF_)
            f.write('BA_DEF_ "BusType" STRING;\n')
            f.write('BA_DEF_ BU_ "NmStationAddress" HEX 0 255;\n')
            f.write('BA_DEF_ BO_ "VFrameFormat" ENUM "StandardCAN","ExtendedCAN","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","reserved","StandardCAN_FD","ExtendedCAN_FD";\n')
            f.write('BA_DEF_ BO_ "GenMsgCycleTime" INT 0 200000;\n')
            f.write('BA_DEF_ BO_ "GenMsgCycleTimeFast" INT 0 200000;\n')
            f.write('BA_DEF_ BO_ "GenMsgStartDelayTime" INT 0 200000;\n')
            f.write('BA_DEF_ BO_ "GenMsgSendType" ENUM "Cyclic","NoMsgSendType";\n')
            f.write('BA_DEF_ BO_ "GenMsgNrOfRepetition" INT 0 100;\n')
            f.write('BA_DEF_ BO_ "GenMsgDelayTime" INT 0 1000;\n')
            f.write('BA_DEF_ BO_ "GenMsgRequestable" INT 0 1;\n')
            f.write('BA_DEF_ BO_ "GenMsgILSupport" ENUM "No","Yes";\n')
            f.write('BA_DEF_ BO_ "NmMessage" ENUM "No","Yes";\n')
            f.write('BA_DEF_ BO_ "DiagRequest" ENUM "No","Yes";\n')
            f.write('BA_DEF_ BO_ "DiagResponse" ENUM "No","Yes";\n')
            f.write('BA_DEF_ SG_ "GenSigSendType" ENUM "NotUsed","OnWrite","OnWriteWithRepetition","OnChange","OnChangeWithRepetition","IfActive","IfActiveWithRepetition","NoSigSendType";\n')
            f.write('BA_DEF_ SG_ "GenSigStartValue" INT 0 100000;\n\n')

            # 6. Default Attribute Values (BA_DEF_DEF_)
            f.write('BA_DEF_DEF_ "BusType" "CAN FD";\n')
            f.write('BA_DEF_DEF_ "NmStationAddress" 0;\n')
            f.write('BA_DEF_DEF_ "VFrameFormat" "StandardCAN_FD";\n')
            f.write('BA_DEF_DEF_ "GenMsgCycleTime" 0;\n')
            f.write('BA_DEF_DEF_ "GenMsgCycleTimeFast" 0;\n')
            f.write('BA_DEF_DEF_ "GenMsgStartDelayTime" 0;\n')
            f.write('BA_DEF_DEF_ "GenMsgSendType" "NoMsgSendType";\n')
            f.write('BA_DEF_DEF_ "GenMsgNrOfRepetition" 0;\n')
            f.write('BA_DEF_DEF_ "GenMsgDelayTime" 0;\n')
            f.write('BA_DEF_DEF_ "GenMsgRequestable" 0;\n')
            f.write('BA_DEF_DEF_ "GenMsgILSupport" "Yes";\n')
            f.write('BA_DEF_DEF_ "NmMessage" "No";\n')
            f.write('BA_DEF_DEF_ "DiagRequest" "No";\n')
            f.write('BA_DEF_DEF_ "DiagResponse" "No";\n')
            f.write('BA_DEF_DEF_ "GenSigSendType" "NoSigSendType";\n')
            f.write('BA_DEF_DEF_ "GenSigStartValue" 0;\n\n')

            # 7. Attribute Values (BA_)
            f.write('BA_ "BusType" "CAN FD";\n')

            # ECU Attributes
            for ecu, addr in self.ecu_addresses.items():
                f.write(f'BA_ "NmStationAddress" BU_ {ecu} {hex_to_dec(addr)};\n')
            
            # Message and Signal Attributes
            print("\nVFrameFormat 설정:")
            for msg_id, msg in self.messages.items():
                # VFrameFormat
                frame_type = msg['frame_type']
                is_extended = msg_id > 0x7FF or (frame_type and 'Extended' in frame_type)
                val = 15 if is_extended else 14
                f.write(f'BA_ "VFrameFormat" BO_ {msg_id} {val};\n')
                type_str = 'Extended' if is_extended else 'Standard'
                print(f"{msg['name']}: CAN FD {type_str} (VFrameFormat={val}, H열='{frame_type}')")

                # Timing
                if msg['cycle_time'] is not None:
                    f.write(f'BA_ "GenMsgCycleTime" BO_ {msg_id} {int(msg["cycle_time"])};\n')
                if msg['cycle_time_fast'] is not None:
                    f.write(f'BA_ "GenMsgCycleTimeFast" BO_ {msg_id} {int(msg["cycle_time_fast"])};\n')
                if msg['start_delay_time'] is not None:
                     f.write(f'BA_ "GenMsgStartDelayTime" BO_ {msg_id} {int(msg["start_delay_time"])};\n')
                
                # Send Type
                send_type_map = {"Cyclic": 0, "NoMsgSendType": 1}
                send_type_val = send_type_map.get(msg['tx_method'])
                if send_type_val is not None:
                    f.write(f'BA_ "GenMsgSendType" BO_ {msg_id} {send_type_val};\n')

                # Repetition
                if msg['nr_of_repetition'] is not None and int(msg['nr_of_repetition']) != 0:
                    f.write(f'BA_ "GenMsgNrOfRepetition" BO_ {msg_id} {int(msg["nr_of_repetition"])};\n')
                
                # Fixed values
                f.write(f'BA_ "GenMsgDelayTime" BO_ {msg_id} 0;\n')
                f.write(f'BA_ "GenMsgRequestable" BO_ {msg_id} 1;\n')

                # Values only if not default
                if str(msg.get('il_support')).lower() == 'no':
                     f.write(f'BA_ "GenMsgILSupport" BO_ {msg_id} 0;\n') # 0=No
                if str(msg.get('nm_message')).lower() == 'yes':
                     f.write(f'BA_ "NmMessage" BO_ {msg_id} 1;\n') # 1=Yes
                if str(msg.get('diag_request')).lower() == 'yes':
                     f.write(f'BA_ "DiagRequest" BO_ {msg_id} 1;\n') # 1=Yes
                if str(msg.get('diag_response')).lower() == 'yes':
                     f.write(f'BA_ "DiagResponse" BO_ {msg_id} 1;\n') # 1=Yes

                # Signal attributes
                for sig_row in msg["signals"]:
                    sig_name = sig_row[13]
                    # GenSigSendType
                    send_type_str = sig_row[27]
                    sig_send_type_map = {"NotUsed":0, "OnWrite":1, "OnWriteWithRepetition":2, "OnChange":3, "OnChangeWithRepetition":4, "IfActive":5, "IfActiveWithRepetition":6, "NoSigSendType":7}
                    sig_send_val = sig_send_type_map.get(send_type_str)
                    if sig_send_val is not None and sig_send_val != 7: # Don't write default
                        f.write(f'BA_ "GenSigSendType" SG_ {msg_id} {sig_name} {sig_send_val};\n')

                    # GenSigStartValue
                    init_val_raw = sig_row[41]
                    if init_val_raw is not None:
                        init_val = hex_to_dec(init_val_raw)
                        f.write(f'BA_ "GenSigStartValue" SG_ {msg_id} {sig_name} {init_val};\n')

        print(f"\nDBC 생성 완료: {output_path}")


if __name__ == "__main__":
    script_path = Path(sys.argv[0])
    # By convention, the output DBC file has the same name as the excel file
    output_dbc_name = Path(EXCEL_FILE).stem + ".dbc"
    output_dbc_path = script_path.parent / output_dbc_name

    generator = CANDBCGenerator(EXCEL_FILE, use_canfd=USE_CANFD)
    generator.parse_signals()
    generator.generate_dbc(output_dbc_path)
    print(f"\n완료! {output_dbc_path}")
