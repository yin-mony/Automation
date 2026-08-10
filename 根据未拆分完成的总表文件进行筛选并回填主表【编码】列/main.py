# -*- coding: utf-8 -*-
"""主表与副表匹配回填的主逻辑文件。"""

from pathlib import Path
import sys

import pandas as pd


class ExcelFile:
    """主表与副表按订单编号匹配，并将 ASIN 回填到主表编码列。"""

    def __init__(
        self,
        path=None,
        totalCol="描述",
        findCol="多年期定价订单编号",
        asinCol="ASIN",
        targetCol="编码（必填）",
        altTargetCol="编码(必填)",
        legacyFindCols=None,
        legacyAsinCols=None,
        separator=",",
    ):
        self.path = Path(path) if path else None
        self.totalCol = totalCol
        self.findCol = findCol
        self.asinCol = asinCol
        self.targetCol = targetCol
        self.altTargetCol = altTargetCol
        self.legacyFindCols = tuple(legacyFindCols or ("myp_order_id",))
        self.legacyAsinCols = tuple(legacyAsinCols or ("asin",))
        self.separator = separator

    def loadData(self, totalPath, subPath):
        """读取主表和单个副表，并清理列名前后空格。"""
        totalDf = pd.read_excel(totalPath)
        subDf = pd.read_excel(subPath)
        totalDf.columns = totalDf.columns.str.strip()
        subDf.columns = subDf.columns.str.strip()
        return totalDf, subDf

    def validateColumns(self, totalDf, subDf, totalCol=None, findCol=None, asinCol=None):
        """校验主表和副表列名，返回实际使用的副表匹配列与 ASIN 列。"""
        totalCol = totalCol or self.totalCol
        findCol = findCol or self.findCol
        asinCol = asinCol or self.asinCol

        if totalCol not in totalDf.columns:
            raise KeyError(f"总表中不存在列: {totalCol}")

        findAliases = self.legacyFindCols if findCol == self.findCol else ()
        asinAliases = self.legacyAsinCols if asinCol == self.asinCol else ()
        resolvedFindCol = self.resolveColumn(subDf, findCol, findAliases, "副表")
        resolvedAsinCol = self.resolveColumn(subDf, asinCol, asinAliases, "副表")
        return resolvedFindCol, resolvedAsinCol

    def resolveColumn(self, df, requestedCol, aliases, tableLabel):
        """按首选列名和旧版别名寻找实际列名。"""
        candidates = list(dict.fromkeys([requestedCol, *aliases]))
        for col in candidates:
            if col in df.columns:
                return col
        expected = " 或 ".join(f"'{col}'" for col in candidates)
        raise KeyError(f"{tableLabel}中不存在列: {expected}")

    def matchAndCollect(self, totalDf, subDf, totalCol=None, findCol=None, asinCol=None):
        """完全匹配主表描述与副表订单编号，并聚合同一订单下的 ASIN。"""
        totalCol = totalCol or self.totalCol
        findCol = findCol or self.findCol
        asinCol = asinCol or self.asinCol
        findCol, asinCol = self.validateColumns(totalDf, subDf, totalCol, findCol, asinCol)

        totalKey = totalDf[totalCol].astype("string").str.strip()
        findKey = subDf[findCol].astype("string").str.strip()

        subResult = subDf.copy()
        subResult[findCol] = findKey
        subResult[asinCol] = subResult[asinCol].astype("string").str.strip()

        totalKeySet = set(totalKey.dropna())
        subResult["is_match"] = findKey.isin(totalKeySet)

        matchDf = subResult.loc[subResult["is_match"], [findCol, asinCol]].copy()
        statDf = (
            matchDf.groupby(findCol, dropna=False)
            .size()
            .reset_index(name="匹配数量")
            .assign(匹配状态="匹配成功")[[findCol, "匹配状态", "匹配数量"]]
        )

        asinMap = (
            matchDf[matchDf[asinCol].notna() & (matchDf[asinCol] != "")]
            .groupby(findCol, dropna=False)[asinCol]
            .apply(self.dedupeJoinAsins)
            .to_dict()
        )

        return {
            "total_key": totalKey,
            "sub_result": subResult,
            "match_df": matchDf,
            "stat_df": statDf,
            "asin_map": asinMap,
            "find_col_used": findCol,
            "asin_col_used": asinCol,
        }

    def fillTargetColumn(self, totalDf, totalKey, asinMap, targetCol=None, mergeExisting=False):
        """将聚合后的 ASIN 写入主表编码列，可选择与已有编码合并去重。"""
        targetCol = targetCol or self.targetCol
        resultDf = totalDf.copy()

        if targetCol in resultDf.columns:
            targetColUsed = targetCol
        elif self.altTargetCol in resultDf.columns:
            targetColUsed = self.altTargetCol
        else:
            targetColUsed = targetCol
            resultDf[targetColUsed] = pd.NA

        mapped = totalKey.map(asinMap)
        if not mergeExisting:
            resultDf[targetColUsed] = mapped
            return resultDf, targetColUsed

        for idx in resultDf.index:
            newValue = mapped.loc[idx]
            if pd.isna(newValue) or (isinstance(newValue, str) and not newValue.strip()):
                continue

            oldValue = resultDf.at[idx, targetColUsed]
            if pd.isna(oldValue) or (isinstance(oldValue, str) and not oldValue.strip()):
                resultDf.at[idx, targetColUsed] = newValue
            else:
                resultDf.at[idx, targetColUsed] = self.mergeCodeStrings(oldValue, newValue)
        return resultDf, targetColUsed

    def printMatchSummary(self, matchResult, totalDfFilled, targetColUsed):
        """打印单个副表的匹配统计和主表回填预览。"""
        subResult = matchResult["sub_result"]
        matchDf = matchResult["match_df"]
        statDf = matchResult["stat_df"]
        findCol = matchResult["find_col_used"]
        asinCol = matchResult["asin_col_used"]

        print(f"总匹配行数: {len(subResult)}")
        print(f"匹配成功: {int(subResult['is_match'].sum())}")
        print(f"匹配失败: {int((~subResult['is_match']).sum())}")
        print(subResult[[findCol, "is_match"]].head(10))

        print(f"\n完全匹配结果（{findCol} + 匹配状态 + 匹配数量）：")
        print(statDf if not statDf.empty else "无匹配成功数据")

        print(f"\n完全匹配结果（{findCol} + 匹配数量 + {asinCol}值）：")
        if matchDf.empty:
            print("无匹配成功数据")
        else:
            grouped = matchDf.groupby(findCol, dropna=False)
            for orderId, group in grouped:
                print(f"{findCol}: {orderId} | 匹配数量: {len(group)}")
                for asinValue in group[asinCol].tolist():
                    print(f"{asinCol}: {asinValue}")

        print("\n主表回填预览（描述 + 编码列）：")
        print(totalDfFilled[[self.totalCol, targetColUsed]].head(10))

    def normalizeSubPaths(self, subPath):
        """将单个副表路径或路径列表统一成 Path 列表。"""
        if subPath is None:
            return []
        if isinstance(subPath, (list, tuple)):
            return [Path(item) for item in subPath if str(item).strip()]
        return [Path(subPath)]

    def run(
        self,
        totalPath=None,
        subPath=None,
        totalCol=None,
        findCol=None,
        asinCol=None,
        targetCol=None,
        printSummary=True,
        saveResult=False,
        outputPath=None,
    ):
        """执行读取、匹配、聚合、回填和可选写回的完整流程。"""
        oldStyleRun = False
        if self.path and subPath is None and totalPath is not None:
            subPath = totalPath
            totalPath = self.path
            oldStyleRun = True
        elif self.path and isinstance(totalPath, (list, tuple)) and subPath is not None:
            outputPath = subPath
            subPath = totalPath
            totalPath = self.path
            oldStyleRun = True

        totalPath = Path(totalPath) if totalPath else self.path
        if not totalPath:
            raise ValueError("请提供主表文件路径。")
        if oldStyleRun:
            saveResult = True

        totalCol = totalCol or self.totalCol
        findCol = findCol or self.findCol
        asinCol = asinCol or self.asinCol
        targetCol = targetCol or self.targetCol

        subPaths = self.normalizeSubPaths(subPath)
        if not subPaths:
            raise ValueError("至少需要一个副表文件路径。")

        for currentPath in subPaths:
            if not currentPath.exists():
                raise FileNotFoundError(f"副表文件不存在: {currentPath}")

        totalDf = pd.read_excel(totalPath)
        totalDf.columns = totalDf.columns.str.strip()
        totalWork = totalDf.copy()

        perSubMatchResults = []
        subResultFrames = []
        matchDfs = []
        statDfs = []

        for index, currentPath in enumerate(subPaths):
            subDf = pd.read_excel(currentPath)
            subDf.columns = subDf.columns.str.strip()
            resolvedFindCol, resolvedAsinCol = self.validateColumns(
                totalWork,
                subDf,
                totalCol,
                findCol,
                asinCol,
            )

            matchResult = self.matchAndCollect(
                totalWork,
                subDf,
                totalCol=totalCol,
                findCol=resolvedFindCol,
                asinCol=resolvedAsinCol,
            )

            totalWork, targetColUsed = self.fillTargetColumn(
                totalWork,
                matchResult["total_key"],
                matchResult["asin_map"],
                targetCol=targetCol,
                mergeExisting=(index > 0),
            )

            entry = {
                **matchResult,
                "sub_path": str(currentPath),
                "total_df_filled_step": totalWork.copy(),
            }
            perSubMatchResults.append(entry)

            subResult = matchResult["sub_result"].copy()
            subResult["副表文件"] = currentPath.name
            subResultFrames.append(subResult)
            matchDfs.append(matchResult["match_df"])
            statDfs.append(matchResult["stat_df"])

            if printSummary:
                print(f"\n========== 副表 {index + 1}/{len(subPaths)}: {currentPath} ==========")
                self.printMatchSummary(matchResult, totalWork, targetColUsed)

        combinedSubResult = (
            pd.concat(subResultFrames, ignore_index=True)
            if len(subResultFrames) > 1
            else subResultFrames[0]
        )
        combinedMatchDf = pd.concat(matchDfs, ignore_index=True) if len(matchDfs) > 1 else matchDfs[0]
        combinedStatDf = pd.concat(statDfs, ignore_index=True) if len(statDfs) > 1 else statDfs[0]
        lastMatchResult = perSubMatchResults[-1]

        result = {
            "total_key": lastMatchResult["total_key"],
            "sub_result": combinedSubResult,
            "match_df": combinedMatchDf,
            "stat_df": combinedStatDf,
            "asin_map": lastMatchResult["asin_map"],
            "total_df_filled": totalWork,
            "target_col_used": targetColUsed,
            "sub_paths": [str(item) for item in subPaths],
            "per_sub_match_results": perSubMatchResults,
        }

        if saveResult:
            savePath = Path(outputPath) if outputPath else totalPath
            savePath.parent.mkdir(parents=True, exist_ok=True)
            try:
                result["total_df_filled"].to_excel(savePath, index=False)
            except PermissionError as exc:
                raise PermissionError(f"无法写入文件: {savePath}。请关闭该Excel文件后重试。") from exc
            result["saved_path"] = str(savePath)
            print(f"\n已写入主表文件: {savePath}")
        return result

    def runInteractive(self):
        """通过命令行交互读取路径并执行回填流程。"""
        totalInput = input("请输入主表路径（必填）: ").strip()
        print("请输入副表路径（必填）。多个文件请逐行输入，输入空行结束：")
        subLines = []
        while True:
            line = input().strip()
            if not line:
                break
            subLines.append(line)

        if not totalInput or not subLines:
            raise ValueError("主表路径和至少一个副表路径都必须输入，不能留空。")

        totalPath = Path(totalInput)
        if not totalPath.exists():
            raise FileNotFoundError(f"主表文件不存在: {totalPath}")

        subPaths = []
        for line in subLines:
            currentPath = Path(line)
            if not currentPath.exists():
                raise FileNotFoundError(f"副表文件不存在: {currentPath}")
            subPaths.append(currentPath)

        outputInput = input("请输入输出主表路径（留空则覆盖原主表）: ").strip()
        outputPath = Path(outputInput) if outputInput else None

        return self.run(
            totalPath=totalPath,
            subPath=subPaths,
            saveResult=True,
            outputPath=outputPath,
        )

    def dedupeJoinAsins(self, values):
        """同一匹配键下多条 ASIN 去重并按首次出现顺序拼接。"""
        ordered = []
        for value in values:
            if pd.isna(value):
                continue
            text = str(value).strip()
            if not text or text.lower() == "nan":
                continue
            ordered.append(text)
        return self.separator.join(list(dict.fromkeys(ordered)))

    def mergeCodeStrings(self, existing, newValue):
        """多副表回填时合并编码字符串，并避免重复追加。"""
        if pd.isna(newValue):
            return existing

        newText = str(newValue).strip()
        if not newText or newText.lower() == "nan":
            return existing

        oldParts = self.splitCodes(existing)
        seen = set(oldParts)
        merged = oldParts + [part for part in self.splitCodes(newText) if part not in seen]
        return self.separator.join(merged)

    def splitCodes(self, value):
        """按英文逗号拆分编码字符串，并去除空片段。"""
        if pd.isna(value):
            return []
        return [part.strip() for part in str(value).split(self.separator) if part.strip()]


if __name__ == "__main__":
    config = {
        "totalPath": r"C:\RPA流程\根据未拆分完成的总表文件进行筛选并回填主表【编码】列\主表.xlsx",
        "subPaths": [
            r"C:\RPA流程\根据未拆分完成的总表文件进行筛选并回填主表【编码】列\副表.xlsx",
        ],
    }

    # 命令行可覆盖：python main.py <主表路径> <副表1> [副表2 ...]
    if len(sys.argv) >= 3:
        config["totalPath"] = sys.argv[1]
        config["subPaths"] = sys.argv[2:]

    if not config["totalPath"].strip():
        raise ValueError("请在 config 中填写 totalPath（主表路径）。")
    if not config["subPaths"]:
        raise ValueError("请在 config 中填写 subPaths（至少一个副表路径）。")

    excel = ExcelFile(config["totalPath"])
    result = excel.run(subPath=config["subPaths"], saveResult=True)
    subResult = result["sub_result"]
    print(f"匹配成功：{int(subResult['is_match'].sum())}，失败：{int((~subResult['is_match']).sum())}")
    print(f"已写入：{result['saved_path']}")
