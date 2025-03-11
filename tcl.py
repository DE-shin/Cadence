import win32com.client
import pythoncom
import pandas as pd
import numpy as np
import os
import logging

logger = logging.getLogger(__name__)

def cadence_translate(dfs):
    try:
        for sheet_name, df in dfs.items():
            df = df.map(lambda x: x.replace(" ", "").replace("\n", ","))
            dfs[sheet_name] = df

        dfs["vrm"][["subnet", "net"]] = dfs["vrm"][["subnet", "net"]].apply(lambda x: x.str.replace(".", "_"))
        dfs["sink"][["subnet", "net"]] = dfs["sink"][["subnet", "net"]].apply(lambda x: x.str.replace(".", "_"))

        dfs["vrm"][["index", "pin"]] = dfs["vrm"][["index", "pin"]].map(lambda x: str(int(float(x))) if x.replace(".", "", 1).isdigit() and float(x).is_integer() else x)
        dfs["sink"][["index", "pin"]] = dfs["sink"][["index", "pin"]].map(lambda x: str(int(float(x))) if x.replace(".", "", 1).isdigit() and float(x).is_integer() else x)
        dfs["disc"][["subnet", "net"]] = dfs["disc"][["subnet", "net"]].apply(lambda x: str(int(float(x))) if x.replace(".", "", 1).isdigit() and float(x).is_integer() else x)
    except Exception as e:
        logger.error(f"Error in cadence_translate: {e}")
    return dfs

def pre_sim(ETL_FILE_PATH, TCL_FOLDER_PATH):
    try:
        logging.info("Excel Initializing...")
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(ETL_FILE_PATH)
    except Exception as e:
        logger.error(f"Error initializing Excel: {e}")
        return None

    try:
        dfs = {}
        for sheet in wb.Sheets:
            data = np.array(sheet.UsedRange.Value)
            try:
                if sheet.Name == "vrm":
                    df = pd.DataFrame(data=data[1:, :6], columns=data[0, :6])
                elif sheet.Name == "sink":
                    df = pd.DataFrame(data=data[1:, :8], columns=data[0, :8])
                elif sheet.Name == "disc":
                    df = pd.DataFrame(data=data[1:, :5], columns=data[0, :5])
                else:
                    continue
                df = df.replace("", np.nan).ffill().astype(str)
                dfs[sheet.Name] = df
            except Exception as e:
                logger.error(f"Error processing sheet {sheet.Name}: {e}")
    except Exception as e:
        logger.error(f"Error reading Excel file: {e}")
    finally:
        logging.info("Excel Closing...")
        wb.Close(False)
        excel.Quit()
        pythoncom.CoUninitialize()

    # Data Pre-processing for each Tool
    dfs = cadence_translate(dfs)
    
    # TCL Script Generation
    # 1. Classify Power/Gnd Net
    nets = set()
    GND = "M_GND"

    for sheet_name, df in dfs.items():
        if "subent" in df.columns():
            df["subnet"].apply(lambda x: nets.update(x.split(",")))
        if "net" in df.columns():
            df["net"].apply(lambda x: nets.update(x.split(",")))

    classify_tcl_commands = [
        "sigrity::clear\n"
        "sigrity::cls\n"
        "set error_nets {}\n\n"
    ]
    for net in nets:
        classify_tcl_commands.append(f" if {{[catch {{sigrity::move net {{PowerNets}} {{{net}}} {{!}}}}]}} {{\n lappend error_nets {{{net}}}\n}}\n")
        classify_tcl_commands.append(f"catch {{sigrity::update net {{PowerGndPair}} {{{GND}}} {{{net}}} {{!}}}}\n")
    classify_tcl_commands.extend([
        "puts \"-----------------------------------------\"\n",
        "puts \"Error Nets : $error_nets\"\n",
        "puts \"-----------------------------------------\"\n"
    ])

    # 2. add VRM, SINK, Discrete
    add_tcl_commands = [
        "sigrity::clear\n"
        "sigrity::cls\n"
        "set error_VRM {}\n"
        "set error_SINK {}\n"
        "set error_DISC {}\n\n"
    ]

    for _, row in dfs["vrm"][["refdes", "net", "pin", "voltage"]].iterrows():
        refdes = row["refdes"]
        port = f"VRM_{refdes}_{row["net"]}"
        pins = row["pin"].split(",")
        v = row["voltage"]

        add_tcl_commands.append(f"catch {{sigrity::add pdcVRM -m -name {{{port}}} -voltage {{{v}}} {{!}}}}\n")
        add_tcl_commands.append(f"catch {{sigrity::link pdcElem {{{port}}} {{Negative Pin}} {{-Circuit {{{refdes}}} -Net {{{GND}}}}} -LinkCktNode {{!}}}}\n")
        for pin in pins:
            add_tcl_commands.extend([
                f"if {{\n",
                f"    [catch {{sigrity::link pdcElem {{{port}}} {{Positive Pin}} {{-Circuit {{{refdes}}} -Node {{{pin}}}}} -LinkCktNode {{!}}}}]\n",
                f"}} {{\n",
                f"    lappend error_VRM {{{refdes}: {pin}}}\n",
                f"}}\n"
            ])

    for _, row in df["sink"][["refdes", "net", "pin", "voltage", "current"]].iterrows():
        refdes = row["refdes"]
        port = f"SINK_{refdes}_{row["net"]}"
        pins = row["pin"].split(",")
        v = row["voltage"]
        i = row["current"]

        if i == "-":
            continue # 소모전류 값이 없으면 pass

        add_tcl_commands.append(f"catch {{sigrity::add pdcSINK -m -name {{{port}}} -current {{{i}}} -lt {{5,%}} -ut {{5,%}} -model {{Equal Current}} {{!}}}}\n")
        add_tcl_commands.append(f"catch {{sigrity::link pdcElem {{{port}}} {{Negative Pin}} {{-Circuit {{{refdes}}} -Net {{{GND}}}}} -LinkCktNode {{!}}}}\n")
        for pin in pins:
            add_tcl_commands.extend([
                f"if {{\n",
                f"    [catch {{sigrity::link pdcElem {{{port}}} {{Positive Pin}} {{-Circuit {{{refdes}}} -Node {{{pin}}}}} -LinkCktNode {{!}}}}]\n",
                f"}} {{\n",
                f"    lappend error_SINK {{{refdes}: {pin}}}\n",
                f"}}\n"
            ])

    for _, row in df["disc"][["refdes", "resistance"]].iterrows():
        refdes = row["refdes"]
        r = row["resistance"]

        add_tcl_commands.extend([
            f"if {{\n",
                f"    [catch {{sigrity::add pdcInter -auto -ckt {{{refdes}}} -resistance {{{r}}} {{!}}}}]\n",
                f"}} {{\n",
                f"    lappend error_DISC {{{refdes}}}\n",
                f"}}\n"
        ])

    add_tcl_commands.extend([
        "puts \"-----------------------------------------\"\n",
        "puts \"Error VRMs : $error_VRM\"\n",
        "puts \"-----------------------------------------\"\n",
        "puts \"Error SINKs : $error_SINK\"\n",
        "puts \"-----------------------------------------\"\n",
        "puts \"Error DISCs : $error_DISC\"\n",
        "puts \"-----------------------------------------\"\n"
    ])

    try:
        logging.info("Saving TCLs...")
        with open(os.path.normpath(os.path.join(TCL_FOLDER_PATH, "classify.tcl")), "w") as f:
            f.writelines("Sample TCL content for classify.tcl")
        with open(os.path.normpath(os.path.join(TCL_FOLDER_PATH, "add.tcl")), "w") as f:
            f.writelines("Sample TCL content for add.tcl")
        logging.info("Saved TCLs")
    except Exception as e:
        logger.error(f"Error saving TCL scripts: {e}")

    return None