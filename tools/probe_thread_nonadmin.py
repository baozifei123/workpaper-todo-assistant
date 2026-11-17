# -*- coding: utf-8 -*-
"""在普通权限下测试不同线程 COM 初始化方式读取 NativeOM。结果写入 JSON。"""
import json
import os
import threading

import comtypes

from workpaper_todo import capture, nativeom

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "..", "testdata", "thread_probe_result.json")
OUT = os.path.abspath(OUT)


def main():
    target = None
    for sw in nativeom.iter_sheet_windows():
        if sw.source_app in ("wps", "excel"):
            target = sw.xlmain
            break
    result = {"target": target}
    if not target:
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(result, f)
        return

    def run(mode):
        def fn():
            init_hr = None
            try:
                if mode == "none":
                    pass
                elif mode == "coinit":
                    init_hr = comtypes.CoInitialize()
                elif mode == "sta":
                    init_hr = comtypes.CoInitializeEx(comtypes.COINIT_APARTMENTTHREADED)
                elif mode == "mta":
                    init_hr = comtypes.CoInitializeEx(comtypes.COINIT_MULTITHREADED)
            except Exception as exc:
                result[mode] = ["INIT_ERR", repr(exc)]
                return
            try:
                info = capture.read_selection(target)
                result[mode] = ["OK", info.file_name, info.cell_address,
                                "init_hr=%r" % (init_hr,)]
            except Exception as exc:
                result[mode] = ["ERR", repr(exc), "init_hr=%r" % (init_hr,)]

        t = threading.Thread(target=fn)
        t.start()
        t.join(15)
        if t.is_alive():
            result[mode] = ["TIMEOUT"]

    for mode in ("none", "coinit", "sta", "mta"):
        run(mode)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
