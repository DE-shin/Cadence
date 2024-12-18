import os
import pandas as pd

# Constant Variables
NETGND = "M_GND"

def classify_power_net(dfs , prj_path) -> None:
    """
    dfs : {sheet_name : dataframe}
    prj_path : cadence project path for save commands
    -----------------------------------------------------------
    sigrity::move net {PowerNets} {$Net} {!} => P/G net classify
    sigrity::update net {PowerGndPair} {$GroundNet} {$Net} {!} => P/G pair
    """

    # Initial Setting
    df_net = pd.concat([dfs["VRM_List"].loc[:, ["NET", "POWER_NET"]], 
                        dfs["SINK_List"].loc[:, ["NET", "POWER_NET"]]])

    power_net = set()

    # Net Extraction
    for col in ["NET", "POWER_NET"]:
        for row in df_net[col]:
            if pd.notna(row):
                values = str(row).replace(".", "_").split("\n")
                values = [v.strip() for v in values]
                power_net.update(values)

    power_net = sorted(list(power_net))

    # TCL Command Generation
    tcl_commands = ["set non_existing_nets {}\n"]

    for net in power_net:
        tcl_commands.append(f"if {{[catch {{sigrity::move net {{PowerNets}} {{{net}}} {{!}}}}]}} {{\n lappend non_existing_nets {{{net}}}\n}}\n")
    
    for net in power_net:
        tcl_commands.append(f"catch {{sigrity::update net {{PowerGndPair}} {{{NETGND}}} {{{net}}} {{!}}}}\n")

    tcl_commands.append("puts \"### ### ### ###\"\n")
    tcl_commands.append("puts \"Non Existing Nets : $non_existing_net\"\n")

    # Save TCL Command
    tcl_script_path = os.path.join(prj_path, "Scripts", "net_classify_power.tcl")
    os.makedirs(os.path.dirname(tcl_script_path), exist_ok=True)

    with open(tcl_script_path, "w") as f:
        f.writelines(tcl_commands)

    print(f"TCL Scripts Generated to {tcl_script_path}")

    return None

def add_VRM_SINK(dfs, prj_path) -> None:

    # Inintial Setting
    df_vrm = dfs["VRM_List"]
    df_sink = dfs["SINK_List"]

    df_vrm["REF"] = df_vrm["REF"].fillna(method="ffill")
    df_sink["REF"] = df_sink["REF"].fillna(method="ffill")

    # TCL Header Generation
    tcl_commands = ["# TCL Script for VRM & SINK Configuration\n"]
    tcl_commands.append("set non_existing_vrms {}\n")
    tcl_commands.append("set non_exisiting_vrm_pins {}\n")
    tcl_commands.append("set non_existing_sinks {}\n")
    tcl_commands.append("set non_exisiting_sink_pins {}\n")

    # TCL Body Generation
    # For VRM
    for idx, row in df_vrm.iterrows():
        refdes = str(row["REF"]).strip()
        nets = str(row["NET"]).replace(".", "_").strip()
        voltage = str(row["VOLTAGE[V]"]).strip()
        pins = str(row["PIN_INDEX"]).strip()

        # Single Net
        if "\n" not in nets:
            command = f"if {{[catch {{sigrity::add pdcVRM -auto -ckt {{{refdes}}} -net {{{nets},{NETGND}}} -voltage {{{voltage}}} {{!}}}}]}} {{\n"
            command += f"    lappend non_existing_vrms {{{refdes}}}\n}}\n"
            tcl_commands.append(command)
        # Multiple Net
        else:
            first_net = nets.split("\n")[0].strip()
            pin_list = [pin.strip() for pin in pins.replace("\n", ",").split(",") if pin.strip()]

            command = f"if {{[catch {{sigrity::add pdcVRM -auto -ckt {{{refdes}}} -net {{{first_net},{NETGND}}} -voltage {{{voltage}}} {{!}}}}]}} {{\n"
            command += f"    lappend non_existing_vrms {{{refdes}}}\n}}\n"
            tcl_commands.append(command)

            for pin in pin_list:
                if pin.isdecimal():
                    pin = str(int(pin))
                command = f"if {{[catch {{sigrity::link pdcElem {{VRM_{refdes}_{first_net}_{NETGND}}} {{Positive Pin}}   {{-Circuit {{{refdes}}} -Node {{{pin}}}}} -LinkCktNode {{!}}}}]}} {{\n"
                command += f"    lappend non_existing_vrm_pins {{{refdes}_{pin}}}\n}}\n"
                tcl_commands.append(command)
    
    # For SINK
    for idx, row in df_sink.iterrows():
        refdes = str(row["REF"]).strip()
        nets = str(row["NET"]).replace(".", "_").strip()
        current = str(row["CURRENT[A]"]).strip()
        pins = str(row["PIN_INDEX"]).strip()

        # Single Net
        if "\n" not in nets:
            command = f"if {{[catch {{sigrity::add pdcSink -auto -ckt {{{refdes}}} -net {{{nets},{NETGND}}} -model {{Equal Current}} -current {{{current}}} -upperTolerance {{5,%}} -lowerTolerance {{5,%}} {{!}}}}]}} {{\n"
            command += f"    lappend non_existing_sinks {{{refdes}}}\n}}\n"
            tcl_commands.append(command)
        # Multiple Net
        else:
            pin_list = [pin.strip() for pin in pins.replace("\n", ",").split(",") if pin.strip()]
            positive_pins = " ".join([f"{{{str(int(pin)) if pin.isdecimal() else pin}}}" for pin in pin_list])

            command = f"if {{[catch {{sigrity::add pdcSink -auto -ckt {{{refdes}}} -net {{{NETGND},{NETGND}}} -positivePin {positive_pins} -model {{Equal Current}} -current {{{current}}} -upperTolerance {{5,%}} -lowerTolerance {{5,%}} {{!}}}}]}} {{\n"
            command += f"    lappend non_existing_sink_pins {{{refdes}_{positive_pins}}}\n}}\n"
            tcl_commands.append(command)

    # TCL Trailer
    tcl_commands.append("puts \"### ### ### ###\"\n")
    tcl_commands.append("puts \"Non Existing VRMs: $non_existing_vrms\"\n")
    tcl_commands.append("puts \"Non Existing Pins: $non_existing_vrm_pins\"\n")
    tcl_commands.append("puts \"### ### ### ###\"\n")
    tcl_commands.append("puts \"Non Existing VRMs: $non_existing_sinks\"\n")
    tcl_commands.append("puts \"Non Existing Pins: $non_existing_sink_pins\"\n")
    tcl_commands.append("puts \"### ### ### ###\"\n")

    # Save TCL Command
    tcl_script_path = os.path.join(prj_path, "Scripts", "add_VRM_SINK.tcl")
    os.makedirs(os.path.dirname(tcl_script_path), exist_ok=True)

    with open(tcl_script_path, "w") as f:
        f.writelines(tcl_commands)

    print(f"TCL Script Generated to {tcl_script_path}")
    
    return None

def add_DCR(dfs, prj_path) -> None:
    """
    dfs : {sheet_name : dataframe}
    prj_path : cadence project path for save commands
    -----------------------------------------------------------
    sigrity::add pdcInter -auto -ckt {$REFDES} -net {$PowerNet, $PowerNet} -resistance {$DCresistance} {!} => add dc resistance at component
    """

    return None