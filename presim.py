import xlwings as xw
import pandas as pd
import numpy as np
from pathlib import Path
import logging
import inspect

logger = logging.getLogger()

# ETL 엑셀 파일 읽는 공통 함수
def _read_excel(etl_file_path):
    """
    엑셀 파일을 읽어 각 시트를 DataFrame으로 변환한 후 전처리하여 반환합니다.
    """
    app = xw.App(visible=False)
    wb = app.books.open(etl_file_path)
    dfs = {
        sheet.name: sheet.used_range.options(pd.DataFrame, header=1, index=False).value.apply(lambda col: col.map(str))
        for sheet in wb.sheets
    }
    for sheet_name, df in dfs.items():
        df = (
            df.replace("", np.nan)
            .ffill()
            .astype(str, errors="ignore")
            .replace({" ": "", "\n": ","}, regex=True)
        )
        dfs[sheet_name] = df
    wb.close()
    app.quit()
    return dfs

# PowerDC
class PdcPresim:
    def __init__(self, GND_NAME, ETL_FILE_PATH):
        # 상수
        self.gnd = GND_NAME
        self.etl_file_path = Path(ETL_FILE_PATH)

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

        self.dfs = _read_excel(self.etl_file_path)

        # Cadence net 표기에 맞게 수정
        self.dfs["vrm"][["subnet", "net"]] = self.dfs["vrm"][["subnet", "net"]].apply(lambda x: x.str.replace(".", "_"))
        self.dfs["vrm"][["index", "pin"]] = self.dfs["vrm"][["index", "pin"]].map(
            lambda x: str(int(float(x))) if x.replace(".", "", 1).isdigit() and float(x).is_integer() else x)

        self.dfs["sink"][["subnet", "net"]] = self.dfs["sink"][["subnet", "net"]].apply(lambda x: x.str.replace(".", "_"))
        self.dfs["sink"][["index", "pin"]] = self.dfs["sink"][["index", "pin"]].map(
            lambda x: str(int(float(x))) if x.replace(".", "", 1).isdigit() and float(x).is_integer() else x)

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
            self.classify_tcl_commands.append(
                f" if {{[catch {{sigrity::move net {{PowerNets}} {{{net}}} {{!}}}}]}} {{\n lappend error_nets {{{net}}}\n}}\n"
            )
            self.classify_tcl_commands.append(
                f"catch {{sigrity::update net {{PowerGndPair}} {{{self.gnd}}} {{{net}}} {{!}}}}\n"
            )
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
                f"catch {{sigrity::add pdcVRM -m -name {{{vrm_port}}} -voltage {{{vrm_v}}} {{!}}}}\n"
            )
            self.add_tcl_commands.append(
                f"catch {{sigrity::link pdcElem {{{vrm_port}}} {{Negative Pin}} {{-Circuit {{{vrm_refdes}}} -Net {{{self.gnd}}}}} -LinkCktNode {{!}}}}\n"
            )
            for vrm_pin in vrm_pins:
                self.add_tcl_commands.extend([
                    f"if {{\n",
                    "    [catch {sigrity::link pdcElem " +
                    f"{{{vrm_port}}} {{Positive Pin}} {{-Circuit {{{vrm_refdes}}} -Node {{{vrm_pin}}}}} -LinkCktNode {{!}}}}]\n",
                    "} {\n",
                    f"    lappend error_vrms {{{vrm_refdes}: {vrm_pin}}}\n",
                    "}\n"
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
            f"catch {{sigrity::add pdcSINK -m -name {{{sink_port}}} -current {{{sink_i}}} -lt {{5,%}} -ut {{5,%}} -model {{Equal Current}} {{!}}}}\n"
        )
        self.add_tcl_commands.append(
            f"catch {{sigrity::link pdcElem {{{sink_port}}} {{Negative Pin}} {{-Circuit {{{sink_refdes}}} -Net {{{self.gnd}}}}} -LinkCktNode {{!}}}}\n"
        )
        for sink_pin in sink_pins:
            self.add_tcl_commands.extend([
                "if {\n",
                "    [catch {sigrity::link pdcElem " +
                f"{{{sink_port}}} {{Positive Pin}} {{-Circuit {{{sink_refdes}}} -Node {{{sink_pin}}}}} -LinkCktNode {{!}}}]\n",
                "} {\n",
                f"    lappend error_SINK {{{sink_refdes}: {sink_pin}}}\n",
                "}\n"
            ])

        # disc sheet
        for _, row in self.dfs["disc"].iterrows():
            disc_refdes = row["refdes"]
            disc_r = row["resistance"]

            self.add_tcl_commands.extend([
                "if {\n",
                f"    [catch {{sigrity::add pdcInter -auto -ckt {{{disc_refdes}}} -resistance {{{disc_r}}} {{!}}}}]\n",
                "} {\n",
                f"    lappend error_DISC {{{disc_refdes}}}\n",
                "}\n"
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
class PsiPresim:
    def __init__(self, GND_NAME, ETL_FILE_PATH):
        # 상수
        self.gnd = GND_NAME
        self.etl_file_path = Path(ETL_FILE_PATH)

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

        self.dfs = _read_excel(self.etl_file_path)

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
            self.classify_tcl_commands.append(
                f" if {{[catch {{sigrity::move net {{PowerNets}} {{{net}}} {{!}}}}]}} {{\n lappend error_nets {{{net}}}\n}}\n"
            )
            self.classify_tcl_commands.append(
                f"catch {{sigrity::update net {{PowerGndPair}} {{{self.gnd}}} {{{net}}} {{!}}}}\n"
            )
        self.classify_tcl_commands.extend([
            "sigrity::save {¡}\n",
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
                "if  {\n",
                f"    [catch {{sigrity::add port -name {{{vrm_port}}} {{¡}}}}]\n",
                "} {\n",
                f"    lappend error_ports {{{vrm_port}}}\n",
                "}\n",
                f"catch {{sigrity::update -name {{{vrm_port}}} -refZ {{1}} {{¡}}}}\n"
            ])

            for pp in vrm_pps.split(","):
                self.add_tcl_commands.extend([
                    "if  {\n",
                    f"    [catch {{sigrity::hook port -name {{{vrm_port}}} -c {{{vrm_refdes}}} -pn {{{pp}}} {{¡}}}}]\n",
                    "} {\n",
                    f"    lappend error_pins {{{pp}}}\n",
                    "}\n"
                ])

            for np_val in vrm_nps.split(","):
                self.add_tcl_commands.extend([
                    "if  {\n",
                    f"    [catch {{sigrity::hook port -name {{{vrm_port}}} -c {{{vrm_refdes}}} -nn {{{np_val}}} {{¡}}}}]\n",
                    "} {\n",
                    f"    lappend error_pins {{{np_val}}}\n",
                    "}\n"
                ])

        # sink sheet
        for _, row in self.dfs["sink"].iterrows():
            sink_refdes = row["refdes"]
            sink_net = row["net"]
            sink_pps = row["pp"]
            sink_nps = row["np"]
            sink_port = f"VRM_{sink_refdes}_{sink_net}" + (
                f"_P{int(float(row['port']))}" if pd.notna(row["port"]) else ""
            )

            self.add_tcl_commands.extend([
                "if  {\n",
                f"    [catch {{sigrity::add port -name {{{sink_port}}} {{¡}}}}]\n",
                "} {\n",
                f"    lappend error_ports {{{sink_port}}}\n",
                "}\n",
                f"catch {{sigrity::update -name {{{sink_port}}} -refZ {{1}} {{¡}}}}\n"
            ])

            for pp in sink_pps.split(","):
                self.add_tcl_commands.extend([
                    "if  {\n",
                    f"    [catch {{sigrity::hook port -name {{{sink_port}}} -c {{{sink_refdes}}} -pn {{{pp}}} {{¡}}}}]\n",
                    "} {\n",
                    f"    lappend error_pins {{{pp}}}\n",
                    "}\n"
                ])

            for np in sink_nps.split(","):
                self.add_tcl_commands.extend([
                    "if  {\n",
                    f"    [catch {{sigrity::hook port -name {{{sink_port}}} -c {{{sink_refdes}}} -nn {{{np}}} {{!}}}}]\n",
                    "} {\n",
                    f"    lappend error_pins {{{np}}}\n",
                    "}\n"
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
                "if  {\n",
                f"    [catch {{sigrity::update circuit -model {{disable}} {{{nc_comp}}} {{¡}}}}]\n",
                "} {\n",
                f"    lappend error_components {{{nc_comp}}}\n",
                "}\n"
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