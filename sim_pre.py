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
        tcl_commands.append(f"catch {{sigrity::update net {{PowerGndPair}} {{NETGND}} {{{net}}} {{!}}}}\n")

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
    

    return None