import xlwings as xw
import pandas as pd
import numpy as np
import os
import logging
import inspect

logger = logging.getLogger()

# PowerDC
class pdc_presim:
    def __init__(self, GND_NAME, ETL_FILE_PATH):
        # 상수
        self.gnd = GND_NAME
        self.etl_file_path = ETL_FILE_PATH

        # 변수
        self.dfs = dict()
        self.classify_tcl_commands = list()
        self.add_tcl_commands = list()
        self.simulation_setup_tcl_commands = list()

        # 함수
        self.initialize()
        self.generate_classify_tcl()
        self.generate_add_tcl()
        self.generate_simulation_setup_tcl()

    def initialize(self):
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 실행 중")

        # 엑셀 파일 읽기 (xlwings)
        app = xw.App(visible=False)
        wb = app.books.open(self.etl_file_path)

        for sheet in wb.Sheets:
            data = np.array(sheet.UsedRange.Value)

            if sheet.name == "vrm":
                df = pd.DataFrame(data=data[1:, :6], columns=data[0, :6])
            elif sheet.name == "sink":
                df = pd.DataFrame(data=data[1:, :8], columns=data[0, :8])
            elif sheet.name == "disc":
                df = pd.DataFrame(data=data[1:, :5], columns=data[0, :5])
            else:
                pass

            df = df.replace("", np.nan).ffill().astype(str)
            self.dfs[sheet.name] = df

        wb.close()
        app.quit()

        # Cadence net 표기에 맞게 수정
        for sheet_name, df in self.dfs.items():
            df = df.map(lambda x: x.replace(" ", "").replace("\n", ","))
            self.dfs[sheet_name] = df

        self.dfs["vrm"][["subnet", "net"]] = self.dfs["vrm"][["subnet", "net"]].apply(lambda x: x.str.replace(".", "_"))
        self.dfs["sink"][["subnet", "net"]] = self.dfs["sink"][["subnet", "net"]].apply(lambda x: x.str.replace(".", "_"))
        self.dfs["vrm"][["index", "pin"]] = self.dfs["vrm"][["index", "pin"]].map(lambda x: str(int(float(x))) if x.replace(".", "", 1).isdigit() and float(x).is_integer() else x)
        self.dfs["sink"][["index", "pin"]] = self.dfs["sink"][["index", "pin"]].map(lambda x: str(int(float(x))) if x.replace(".", "", 1).isdigit() and float(x).is_integer() else x)

        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 완료")
        return None

    def generate_classify_tcl(self):
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 실행 중")

        nets = set()

        self.classify_tcl_commands.extend([
            "sigrity::clear\n",
            "sigrity::cls\n",
            "set error_nets {}\n\n"
        ])

        for sheet_name, df in self.dfs.items():
            if "net" in df.columns:
                df["net"].apply(lambda x: nets.update(x.split(",")))
            if "subnet" in df.columns:
                df["subnet"].apply(lambda x: nets.update(x.split(",")))

        for net in nets:
            self.classify_tcl_commands.append(f" if {{[catch {{sigrity::move net {{PowerNets}} {{{net}}} {{!}}}}]}} {{\n lappend error_nets {{{net}}}\n}}\n")
            self.classify_tcl_commands.append(f"catch {{sigrity::update net {{PowerGndPair}} {{{self.gnd}}} {{{net}}} {{!}}}}\n")
        self.classify_tcl_commands.extend([
            f"sigrity::save {{!}}\n",
            "puts \n\"=============================================\"\n",
            "puts \"Error Nets : $error_nets\"\n",
            "puts \n\"=============================================\"\n"
        ])

        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 완료")
        return None

    def generate_add_tcl(self):
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 실행 중")

        self.add_tcl_commands.extend([
            "sigrity::clear\n",
            "sigrity::cls\n",
            "set error_vrms {}\n",
            "set error_sinks {}\n",
            "set error_discs {}\n\n"
        ])

        # vrm sheet
        for _, row in self.dfs["vrm"].iterrows():
            vrm_refdes = row["refdes"]
            vrm_net = row["net"]
            vrm_pins = row["pin"].split(",")
            vrm_v = row["v"]
            vrm_port = f"VRM_{vrm_refdes}_{vrm_net}"

            self.add_tcl_commands.append(
                f"catch {{sigrity::add pdcVRM -m -name {{{vrm_port}}} -voltage {{{vrm_v}}} {{!}}}}\n")
            self.add_tcl_commands.append(
                f"catch {{sigrity::link pdcElem {{{vrm_port}}} {{Negative Pin}} {{-Circuit {{{vrm_refdes}}} -Net {{{self.gnd}}}}} -LinkCktNode {{!}}}}\n")
            for vrm_pin in vrm_pins:
                self.add_tcl_commands.extend([
                    f"if {{\n",
                    f"    [catch {{sigrity::link pdcElem {{{vrm_port}}} {{Positive Pin}} {{-Circuit {{{vrm_refdes}}} -Node {{{vrm_pin}}}}} -LinkCktNode {{!}}}}]\n",
                    f"}} {{\n",
                    f"    lappend error_vrms {{{vrm_refdes}: {vrm_pin}}}\n",
                    f"}}\n"
                ])

        # sink sheet
        for _, row in self.dfs["sink"].iterrows():
            sink_refdes = row["refdes"]
            sink_net = row["net"]
            sink_pins = row["pin"].split(",")
            sink_v = row["voltage"]
            sink_i = row["current"]
            sink_port = f"SINK_{sink_refdes}_{sink_net}"

        self.add_tcl_commands.append(
            f"catch {{sigrity::add pdcSINK -m -name {{{sink_port}}} -current {{{sink_i}}} -lt {{5,%}} -ut {{5,%}} -model {{Equal Current}} {{!}}}}\n")
        self.add_tcl_commands.append(
            f"catch {{sigrity::link pdcElem {{{sink_port}}} {{Negative Pin}} {{-Circuit {{{sink_refdes}}} -Net {{{self.gnd}}}}} -LinkCktNode {{!}}}}\n")
        for sink_pin in sink_pins:
            self.add_tcl_commands.extend([
                f"if {{\n",
                f"    [catch {{sigrity::link pdcElem {{{sink_port}}} {{Positive Pin}} {{-Circuit {{{sink_refdes}}} -Node {{{sink_pin}}}}} -LinkCktNode {{!}}}}]\n",
                f"}} {{\n",
                f"    lappend error_SINK {{{sink_refdes}: {sink_pin}}}\n",
                f"}}\n"
            ])

        # disc sheet
        for _, row in self.dfs["disc"].iterrows():
            disc_refdes = row["refdes"]
            disc_r = row["resistance"]

            self.add_tcl_commands.extend([
                f"if {{\n",
                f"    [catch {{sigrity::add pdcInter -auto -ckt {{{disc_refdes}}} -resistance {{{disc_r}}} {{!}}}}]\n",
                f"}} {{\n",
                f"    lappend error_DISC {{{disc_refdes}}}\n",
                f"}}\n"
            ])

        self.classify_tcl_commands.extend([
            f"sigrity::save {{!}}\n",
            "puts \n\"=============================================\"\n",
            "puts \"Error VRMs : $error_vrms\"\n",
            "puts \"Error SINKs : $error_sinks\"\n",
            "puts \"Error DISCs : $error_discs\"\n",
            "puts \n\"=============================================\"\n"
        ])

        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 완료")
        return None

    def generate_simulation_setup_tcl(self):
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 실행 중")

        self.simulation_setup_tcl_commands.extend([
            "sigrity::update option -MaxEdgeLength {0.001000} {!}\n",
            "sigrity::save {!}\n",
            # start simulation
        ])

        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 완료")
        return None


# PowerSI
class psi_presim:
    def __init__(self, GND_NAME, ETL_FILE_PATH):
        # 상수
        self.gnd = GND_NAME
        self.etl_file_path = ETL_FILE_PATH

        # 변수
        self.dfs = dict()
        self.classify_tcl_commands = list()
        self.add_tcl_commands = list()
        self.nc_tcl_commands = list()
        self.assign_tcl_commands = list()

        # 함수
        self.initialize()
        self.generate_classify_tcl()
        self.generate_add_tcl()
        self.assign_tcl_commands()

    def initialize(self):
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 실행 중")

        # 엑셀 파일 읽기 (xlwings)
        app = xw.App(visible=False)
        wb = app.books.open(self.etl_file_path)
        self.dfs = {sheet.name: sheet.used_range.options(pd.DataFrame, header=1, index=False).value.apply(lambda x: x.map(str)) for sheet in wb.sheets}
        wb.close()
        app.quit()

        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 완료")
        return None

    def generate_classify_tcl(self):
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 실행 중")

        nets = set()

        self.classify_tcl_commands.extend([
            "sigrity::clear\n",
            "sigrity::cls\n",
            "set error_nets {}\n\n"
        ])

        for sheet_name, df in self.dfs.items():
            if "net" in df.columns:
                df["net"].apply(lambda x: nets.update(x.split(",")))

        for net in nets:
            self.classify_tcl_commands.append(f" if {{[catch {{sigrity::move net {{PowerNets}} {{{net}}} {{!}}}}]}} {{\n lappend error_nets {{{net}}}\n}}\n")
            self.classify_tcl_commands.append(f"catch {{sigrity::update net {{PowerGndPair}} {{{self.gnd}}} {{{net}}} {{!}}}}\n")
        self.classify_tcl_commands.extend([
            f"sigrity::save {{!}}\n",
            "puts \n\"=============================================\"\n",
            "puts \"Error Nets : $error_nets\"\n",
            "puts \n\"=============================================\"\n"
        ])

        print(nets)
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 완료")
        return None

    def generate_add_tcl(self):
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 실행 중")

        self.add_tcl_commands.extend([
            "sigrity::clear\n",
            "sigrity::cls\n",
            "set error_ports {}\n",
            "set error_pins {}\n",
        ])

        # vrm sheet
        for _, row in self.dfs["vrm"].iterrows():
            vrm_refdes = row["refdes"]
            vrm_net = row["net"]
            vrm_pps = row["pp"]
            vrm_nps = row["np"]
            vrm_port = f"VRM_{vrm_refdes}_{vrm_net}"

            self.add_tcl_commands.extend([
                f"if  {{\n",
                f"    [catch {{sigrity::add port -name {{{vrm_port}}} {{!}}}}]\n",
                f"}} {{\n",
                f"    lappend error_ports {{{vrm_port}}}\n",
                f"}}\n",
                f"catch {{sigrity::update -name {{{vrm_port}}} -refZ {{1}} {{!}}}}\n"
            ])

            for pp in vrm_pps.split(","):
                self.add_tcl_commands.extend([
                    f"if  {{\n",
                    f"    [catch {{sigrity::hook port -name {{{vrm_port}}} -c {{{vrm_refdes}}} -pn {{{pp}}} {{!}}}}]\n",
                    f"}} {{\n",
                    f"    lappend error_pins {{{pp}}}\n",
                    f"}}\n"
                ])

            for np in vrm_nps.split(","):
                self.add_tcl_commands.extend([
                    f"if  {{\n",
                    f"    [catch {{sigrity::hook port -name {{{vrm_port}}} -c {{{vrm_refdes}}} -nn {{{np}}} {{!}}}}]\n",
                    f"}} {{\n",
                    f"    lappend error_pins {{{np}}}\n",
                    f"}}\n"
                ])

        # sink sheet
        for _, row in self.dfs["sink"].iterrows():
            sink_refdes = row["refdes"]
            sink_net = row["net"]
            sink_pps = row["pp"]
            sink_nps = row["np"]
            sink_port = f"VRM_{sink_refdes}_{sink_net}" + (f"_P{int(float(row["port"]))}" if not pd.isna(row["port"]) else "")

            self.add_tcl_commands.extend([
                f"if  {{\n",
                f"    [catch {{sigrity::add port -name {{{sink_port}}} {{!}}}}]\n",
                f"}} {{\n",
                f"    lappend error_ports {{{sink_port}}}\n",
                f"}}\n",
                f"catch {{sigrity::update -name {{{sink_port}}} -refZ {{1}} {{!}}}}\n"
            ])

            for pp in sink_pps.split(","):
                self.add_tcl_commands.extend([
                    f"if  {{\n",
                    f"    [catch {{sigrity::hook port -name {{{sink_port}}} -c {{{sink_refdes}}} -pn {{{pp}}} {{!}}}}]\n",
                    f"}} {{\n",
                    f"    lappend error_pins {{{pp}}}\n",
                    f"}}\n"
                ])

            for np in sink_nps.split(","):
                self.add_tcl_commands.extend([
                    f"if  {{\n",
                    f"    [catch {{sigrity::hook port -name {{{sink_port}}} -c {{{sink_refdes}}} -nn {{{np}}} {{!}}}}]\n",
                    f"}} {{\n",
                    f"    lappend error_pins {{{np}}}\n",
                    f"}}\n"
                ])

        self.classify_tcl_commands.extend([
            f"sigrity::save {{!}}\n",
            "puts \n\"=============================================\"\n",
            "puts \"Error Ports : $error_ports\"\n",
            "puts \"Error Pins : $error_pins\"\n",
            "puts \n\"=============================================\"\n"
        ])

        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 완료")
        return None

    def generate_nc_tcl(self):
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 실행 중")

        self.nc_tcl_commands.extend([
            "sigrity::clear\n",
            "sigrity::cls\n",
            "set error_components {}\n\n"
        ])

        for _, row in self.dfs["nc"].iterrows():
            nc_comp = row["refdes"]

            self.nc_tcl_commands.extend([
                f"if  {{\n",
                f"    [catch {{sigrity::update circuit -model {{disable}} {{{nc_comp}}} {{!}}}}]\n",
                f"}} {{\n",
                f"    lappend error_components {{{nc_comp}}}\n",
                f"}}\n"
            ])

        self.classify_tcl_commands.extend([
            f"sigrity::save {{!}}\n",
            "puts \n\"=============================================\"\n",
            "puts \"Error Components : $error_components\"\n",
            "puts \n\"=============================================\"\n"
        ])

        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 완료")
        return None

    def generate_assign_tcl(self):
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 실행 중")

        """
        sigrity::update cktdef {$p/n} -head {ExtNode = 1 2
        } -Definition {S1 1 2 3 2
        + Model = ""
        V 3 2 0} -check {!}
        """

        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 완료")
        return None