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
    """
    Generate a TCL script for configuring VRM and SINK elements using the provided dataframes.

    Args:
        dfs (dict): A dictionary containing VRM_List and SINK_List dataframes.
        prj_path (str): Project path where the TCL script will be saved.
    """
    # Initial Setting
    df_vrm = dfs["VRM_List"]
    df_sink = dfs["SINK_List"]

    df_vrm["REF"] = df_vrm["REF"].fillna(method="ffill")
    df_sink["REF"] = df_sink["REF"].fillna(method="ffill")

    # TCL Header Generation
    tcl_commands = [
        "# TCL Script for VRM & SINK Configuration\n",
        "set non_existing_vrms {}\n",
        "set non_exisiting_vrm_pins {}\n",
        "set non_existing_sinks {}\n",
        "set non_exisiting_sink_pins {}\n",
    ]

    # TCL Body Generation - For VRM
    for _, row in df_vrm.iterrows():
        refdes = str(row["REF"]).strip()
        nets = str(row["NET"]).replace(".", "_").strip()
        voltage = str(row["VOLTAGE[V]"]).strip()
        pins = str(row["PIN_INDEX"]).strip()

        if "\n" not in nets:  # Single Net
            tcl_commands.append(
                f"if {{[catch {{sigrity::add pdcVRM -auto -ckt {{{refdes}}} \
"
                f"    -net {{{nets},{NETGND}}} -voltage {{{voltage}}} {{!}}}}]}} {{\n"
                f"    lappend non_existing_vrms {{{refdes}}}\n}}
"
            )
        else:  # Multiple Nets
            first_net = nets.split("\n")[0].strip()
            pin_list = [pin.strip() for pin in pins.replace("\n", ",").split(",") if pin.strip()]

            tcl_commands.append(
                f"if {{[catch {{sigrity::add pdcVRM -auto -ckt {{{refdes}}} \
"
                f"    -net {{{first_net},{NETGND}}} -voltage {{{voltage}}} {{!}}}}]}} {{\n"
                f"    lappend non_existing_vrms {{{refdes}}}\n}}
"
            )

            for pin in pin_list:
                pin = str(int(pin)) if pin.isdecimal() else pin
                tcl_commands.append(
                    f"if {{[catch {{sigrity::link pdcElem {{VRM_{refdes}_{first_net}_{NETGND}}} \
"
                    f"    {{Positive Pin}} -Circuit {{{refdes}}} -Node {{{pin}}} -LinkCktNode {{!}}}}]}} {{\n"
                    f"    lappend non_existing_vrm_pins {{{refdes}_{pin}}}\n}}
"
                )

    # TCL Body Generation - For SINK
    for _, row in df_sink.iterrows():
        refdes = str(row["REF"]).strip()
        nets = str(row["NET"]).replace(".", "_").strip()
        current = str(row["CURRENT[A]"]).strip()
        pins = str(row["PIN_INDEX"]).strip()

        if "\n" not in nets:  # Single Net
            tcl_commands.append(
                f"if {{[catch {{sigrity::add pdcSink -auto -ckt {{{refdes}}} \
"
                f"    -net {{{nets},{NETGND}}} -model {{Equal Current}} \
"
                f"    -current {{{current}}} -upperTolerance {{5,%}} \
"
                f"    -lowerTolerance {{5,%}} {{!}}}}]}} {{\n"
                f"    lappend non_existing_sinks {{{refdes}}}\n}}
"
            )
        else:  # Multiple Nets
            pin_list = [pin.strip() for pin in pins.replace("\n", ",").split(",") if pin.strip()]
            positive_pins = " ".join(
                [f"{{{str(int(pin)) if pin.isdecimal() else pin}}}" for pin in pin_list]
            )

            tcl_commands.append(
                f"if {{[catch {{sigrity::add pdcSink -auto -ckt {{{refdes}}} \
"
                f"    -net {{{NETGND},{NETGND}}} -positivePin {positive_pins} \
"
                f"    -model {{Equal Current}} -current {{{current}}} \
"
                f"    -upperTolerance {{5,%}} -lowerTolerance {{5,%}} {{!}}}}]}} {{\n"
                f"    lappend non_existing_sink_pins {{{refdes}_{positive_pins}}}\n}}
"
            )

    # TCL Trailer
    tcl_commands.extend([
        "puts \"### ### ### ###\"\n",
        "puts \"Non Existing VRMs: $non_existing_vrms\"\n",
        "puts \"Non Existing Pins: $non_existing_vrm_pins\"\n",
        "puts \"### ### ### ###\"\n",
        "puts \"Non Existing SINKs: $non_existing_sinks\"\n",
        "puts \"Non Existing Pins: $non_existing_sink_pins\"\n",
        "puts \"### ### ### ###\"\n",
    ])

    # Save TCL Script
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