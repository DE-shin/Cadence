import os
import pandas as pd

# Constant Variables
NETGND = "M_GND"

def net_classify_power(dfs , prj_path) -> None:
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

def port_add(dfs, prj_path):
    """
    dfs : {sheet_name : dataframe}
    prj_path : cadence project path for save commands
    -----------------------------------------------------------
    sigrity::add pdcVRM -auto -ckt {$REFDES} -net {$PowerNet, $GroundNet} -voltage {$Voltage} {!} => add VRM by Net
    sigrity::add pdcVRM -auto -ckt {$REFDES} -positivePin {$PositivPin} -negative {$NegativePin} -voltage {$Voltage} {!} => add VRM by Pin
    
    sigrity::add pdcSink -auto -ckt {$REFDES} -net {$PowerNet, $GroundNet} -model {Equal Current} -current {$Current} {!} => add SINK by Net
    sigrity::add pdcSink -auto -ckt {$REFDES} -positivePin {$PositivePin} -negative {$NegativePin} -model {Equal Current} -current {$Current} {!} => add SINK by Pin
    sigrity::add pdcSink -atuo -ckt {$REFDES} -net{$GroundNet, $GroundNet} -positivePin {$PositivePin} -model {Equal Current} -current {$Current} -upperTolerance {5,%}, -lowerTolerance {5,%} {!}

    sigrity::add pdcInter -auto -ckt {$REFDES} -net {$PowerNet, $PowerNet} -resistance {$DCresistance} {!} => add dc resistance at component
    """

    # Initial Setting
    df_vrm = dfs["VRM_List"]
    df_sink = dfs["SINK_List"]

    df_vrm["REF"] = df_vrm["REF"].fillna(method="ffill")
    df_sink["REF"] = df_sink["REF"].fillna(method="ffill")

    # TCL Command Generation
    tcl_commands = ["# TCL Script for VRM Configuration\n"]
    tcl_commands.append("set non_existing_vrms {}\n")
    non_existing_pins = []

    for idx, row in df_vrm.iterrows():
        refdes = str(row["REF"]).strip()
        nets = str(row["NET"]).replace(".", "_").strip()
        voltage = str(row["VOLTAGE[V]"]).strip()
        pins = str(row["PIN_INDEX"]).strip()

        if "\n" not in nets:  # Single Net
            command = f"if {{[catch {{sigrity::add pdcVRM -auto -ckt {{{refdes}}} -net {{{nets},{NETGND}}} -voltage {{{voltage}}} {{!}}}}]}} {{\n"
            command += f"    lappend non_existing_vrms {{{refdes}}}\n}}\n"
            tcl_commands.append(command)
        else:  # Multiple Nets
            # 첫 번째 net 가져오기
            first_net = nets.split("\n")[0].strip()

            # 첫 번째 명령어 생성 (예외처리)
            command = f"if {{[catch {{sigrity::add pdcVRM -auto -ckt {{{refdes}}} -net {{{nets},{NETGND}}} -voltage {{{voltage}}} {{!}}}}]}} {{\n"
            command += f"    lappend non_existing_vrms {{{refdes}}}\n}}\n"
            tcl_commands.append(command)

            # NETGND 값 가져오기 (기본값 M_GND)
            netgnd = str(row.get("NETGND", "M_GND")).strip()

            # PIN 리스트 생성
            pin_list = [pin.strip() for pin in pins.replace("\n", ",").split(",") if pin.strip()]

            # 두 번째 명령어 생성
            for pin in pin_list:
                try:
                    command = f"sigrity::link pdcElem {{VRM_{refdes}_{first_net}_{netgnd}}} {{PositivePin}}  {{-Circuit {{{refdes}}} -Node {{{pin}}}}} -LinkCktNode {{!}}\n"
                    tcl_commands.append(command)
                except Exception as e:
                    # 오류 발생 시, non_existing_pins에 저장
                    non_existing_pins.append(f"{refdes}_{pin}")

    tcl_commands.append("puts \"### ### ### ###\"\n")
    tcl_commands.append("puts \"Non Existing VRMs: $non_existing_vrms\"\n")
    
    if non_existing_pins:
        tcl_commands.append("puts \"Non Existing Pins: $non_existing_pins\"\n")

    # Save TCL Command
    tcl_script_path = os.path.join(prj_path, "Scripts", "port_add.tcl")
    os.makedirs(os.path.dirname(tcl_script_path), exist_ok=True)

    with open(tcl_script_path, "w") as f:
        f.writelines(tcl_commands)

    print(f"TCL Script Generated to {tcl_script_path}")

    return None