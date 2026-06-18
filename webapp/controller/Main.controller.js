sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/ui/core/ValueState",
    "sap/m/MessageToast"
], (Controller, JSONModel, Filter, FilterOperator, ValueState, MessageToast) => {
    "use strict";

    const PL_TYPES = [
        { key: "SA", name: "매출액" },
        { key: "CO", name: "매출원가" },
        { key: "SG", name: "판매관리비" },
        { key: "OI", name: "영업외수익" },
        { key: "OE", name: "영업외비용" }
    ];

    const AMOUNT_FIELDS = [
        "previous2Amount",
        "previousAmount",
        "currentAmount"
    ];

    return Controller.extend("code.t4.ui5.fi05.controller.Main", {
        onInit() {
            this.getView().setModel(new JSONModel(this._createInitialSearchData()), "search");
            this.getView().setModel(new JSONModel(this._createInitialPlData()), "pl");
        },

        onSearch() {
            const oSearch = this.getView().getModel("search").getData();
            const sBukrs = String(oSearch.Bukrs || "").trim().toUpperCase();
            const sGjahr = String(oSearch.Gjahr || "").trim();
            const sWeeks = this._padWeek(oSearch.Weeks);

            if (!sBukrs || !sGjahr || !sWeeks) {
                MessageToast.show("회사코드, 회계연도, 회계주차를 입력해주세요.");
                return;
            }

            const aPeriods = this._getDisplayPeriods(Number(sGjahr), Number(sWeeks));

            Promise.all(aPeriods.map((oPeriod) => this._readPeriodData(sBukrs, oPeriod)))
                .then((aResults) => {
                    const oPeriodData = {};

                    aPeriods.forEach((oPeriod, iIndex) => {
                        oPeriodData[oPeriod.key] = aResults[iIndex];
                    });

                    const aItems = this._buildStatementItems(oPeriodData);

                    if (!aItems.length) {
                        this.getView().getModel("pl").setData(this._createInitialPlData());
                        MessageToast.show("조회된 결과가 없습니다.");
                        return;
                    }

                    this.getView().getModel("pl").setData({
                        isSearched: true,
                        periods: this._createPeriodMap(aPeriods),
                        items: aItems
                    });

                    this._expandPlTree(1);
                    MessageToast.show("손익계산서 조회가 완료되었습니다.");
                })
                .catch(() => {
                    MessageToast.show("손익계산서 조회 중 오류가 발생했습니다.");
                });
        },

        _expandPlTree(iLevel) {
            const oTree = this.byId("plTree");
            const oBinding = oTree && oTree.getBinding("items");

            if (oBinding && typeof oBinding.expandToLevel === "function") {
                oBinding.expandToLevel(iLevel);
            }
        },

        onReset() {
            this.getView().getModel("search").setData(this._createInitialSearchData());
            this.getView().getModel("pl").setData(this._createInitialPlData());
            MessageToast.show("초기화되었습니다.");
        },

        onExportPdf() {
            if (!this.getView().getModel("pl").getProperty("/isSearched")) {
                return;
            }

            const sHtml = this._buildExportHtml("pdf");
            const oPrintWindow = window.open("", "_blank");

            if (!oPrintWindow) {
                MessageToast.show("팝업이 차단되어 PDF 출력 창을 열 수 없습니다.");
                return;
            }

            oPrintWindow.document.open();
            oPrintWindow.document.write(sHtml);
            oPrintWindow.document.close();
            oPrintWindow.focus();

            setTimeout(() => {
                oPrintWindow.print();
            }, 300);
        },

        onExportExcel() {
            if (!this.getView().getModel("pl").getProperty("/isSearched")) {
                return;
            }

            const sHtml = this._buildExportHtml("excel");
            this._downloadFile("\uFEFF" + sHtml, "손익계산서.xls", "application/vnd.ms-excel;charset=utf-8");
        },

        formatCurrencyAmount(vAmount, sCurrency) {
            const iAmount = Number(vAmount || 0);
            const iDigits = sCurrency === "KRW" ? 0 : 2;
            const sAmount = Math.abs(iAmount).toLocaleString("ko-KR", {
                minimumFractionDigits: iDigits,
                maximumFractionDigits: iDigits
            });

            return iAmount < 0 ? "(" + sAmount + ")" : sAmount;
        },

        formatCurrencyState(vAmount, sRowType) {
            return sRowType === "total" && Number(vAmount || 0) < 0 ? ValueState.Error : ValueState.None;
        },

        _createInitialSearchData() {
            const oCurrentPeriod = this._getRelativeFiscalWeek(new Date(), -1);

            return {
                Bukrs: "",
                Gjahr: String(oCurrentPeriod.year),
                Weeks: this._padWeek(oCurrentPeriod.week)
            };
        },

        _createInitialPlData() {
            return {
                isSearched: false,
                periods: this._createEmptyPeriodMap(),
                items: []
            };
        },

        _createEmptyPeriodMap() {
            return {
                previous2: { title: "", periodText: "" },
                previous: { title: "", periodText: "" },
                current: { title: "", periodText: "" }
            };
        },

        _buildExportHtml(sMode) {
            const oPlData = this.getView().getModel("pl").getData();
            const aColumns = this._getExportPeriodColumns(oPlData.periods);
            const aPeriodLines = this._getExportPeriodLines(oPlData.periods);
            const aRows = this._flattenExportRows(oPlData.items || []);
            const bExcel = sMode === "excel";
            const sTitle = "손익계산서";
            const sTableBorder = bExcel ? "0.5pt solid #b7b7b7" : "1px solid #111";

            return [
                "<!DOCTYPE html>",
                "<html>",
                "<head>",
                "<meta charset=\"UTF-8\">",
                bExcel ? "<meta http-equiv=\"Content-Type\" content=\"application/vnd.ms-excel; charset=UTF-8\">" : "",
                "<title>" + sTitle + "</title>",
                "<style>",
                "body{font-family:Arial,'Malgun Gothic',sans-serif;margin:18px 22px;color:#111;}",
                ".exportTop{position:relative;margin-bottom:8px;min-height:28px;font-size:13px;}",
                ".company{position:absolute;left:0;bottom:0;font-weight:700;}",
                ".unit{position:absolute;right:0;bottom:0;font-weight:700;}",
                ".title{text-align:center;font-size:26px;font-weight:700;margin:4px 0 10px;}",
                ".periods{text-align:center;font-size:13px;line-height:1.65;margin-bottom:12px;}",
                "table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:12px;}",
                "th,td{border:" + sTableBorder + ";padding:6px 8px;vertical-align:middle;}",
                "th{background:#f3f3f3;text-align:center;font-weight:700;}",
                ".account{text-align:left;}",
                ".amount{text-align:right;white-space:nowrap;}",
                ".detail .account{padding-left:24px;}",
                ".subtotal td{border-top:2px solid #666;font-weight:700;}",
                ".total td{border-top:2px solid #333;border-bottom:2px solid #333;font-weight:700;}",
                ".negativeTotal{color:#c00000;}",
                "@media print{@page{size:A4 landscape;margin:12mm;}body{margin:0;}}",
                "</style>",
                "</head>",
                "<body>",
                "<div class=\"exportTop\"><div class=\"company\">(주)누어바라</div><div class=\"unit\">(단위: 원)</div></div>",
                "<div class=\"title\">" + sTitle + "</div>",
                "<div class=\"periods\">" + aPeriodLines.map((sLine) => this._escapeHtml(sLine)).join("<br>") + "</div>",
                "<table>",
                "<colgroup><col style=\"width:40%\">" + aColumns.map(() => "<col style=\"width:20%\">").join("") + "</colgroup>",
                "<thead><tr><th>과목</th>" + aColumns.map((oColumn) => "<th>" + this._escapeHtml(oColumn.title) + "</th>").join("") + "</tr></thead>",
                "<tbody>",
                this._buildExportTableRowsHtml(aRows, aColumns),
                "</tbody>",
                "</table>",
                "</body>",
                "</html>"
            ].join("");
        },

        _getExportPeriodColumns(oPeriods) {
            return [
                { key: "currentAmount", title: oPeriods.current && oPeriods.current.title },
                { key: "previousAmount", title: oPeriods.previous && oPeriods.previous.title },
                { key: "previous2Amount", title: oPeriods.previous2 && oPeriods.previous2.title }
            ].map((oColumn, iIndex) => Object.assign({
                title: oColumn.title || "제 " + (iIndex + 1) + "기"
            }, oColumn));
        },

        _getExportPeriodLines(oPeriods) {
            return [oPeriods.current, oPeriods.previous, oPeriods.previous2]
                .filter((oPeriod) => oPeriod && oPeriod.periodText)
                .map((oPeriod) => oPeriod.periodText);
        },

        _flattenExportRows(aRows, iLevel = 0) {
            return (aRows || []).reduce((aResult, oRow) => {
                aResult.push(Object.assign({ level: iLevel }, oRow));
                return aResult.concat(this._flattenExportRows(oRow.children || [], iLevel + 1));
            }, []);
        },

        _buildExportTableRowsHtml(aRows, aColumns) {
            return aRows.map((oRow) => {
                const sRowClass = [oRow.rowType, oRow.level > 0 ? "detail" : ""].filter(Boolean).join(" ");
                const sAccount = this._escapeHtml(oRow.name || "");
                const sCells = aColumns.map((oColumn) => {
                    const iAmount = Number(oRow[oColumn.key] || 0);
                    const sClass = oRow.rowType === "total" && iAmount < 0 ? "amount negativeTotal" : "amount";

                    return "<td class=\"" + sClass + "\">" + this.formatCurrencyAmount(iAmount, oRow.Waers) + "</td>";
                }).join("");

                return "<tr class=\"" + sRowClass + "\"><td class=\"account\">" + sAccount + "</td>" + sCells + "</tr>";
            }).join("");
        },

        _downloadFile(sContent, sFileName, sMimeType) {
            const oBlob = new Blob([sContent], { type: sMimeType });
            const sUrl = URL.createObjectURL(oBlob);
            const oLink = document.createElement("a");

            oLink.href = sUrl;
            oLink.download = sFileName;
            document.body.appendChild(oLink);
            oLink.click();
            document.body.removeChild(oLink);
            URL.revokeObjectURL(sUrl);
        },

        _escapeHtml(sValue) {
            return String(sValue || "")
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#39;");
        },

        _readPeriodData(sBukrs, oPeriod) {
            return Promise.all([
                this._readEntitySet("/ZCDS_D4_FI_0002", sBukrs, oPeriod),
                this._readEntitySet("/ZCDS_D4_FI_0001", sBukrs, oPeriod)
            ]).then((aResults) => ({
                summary: aResults[0],
                detail: aResults[1]
            }));
        },

        _readEntitySet(sEntitySet, sBukrs, oPeriod) {
            const oModel = this.getOwnerComponent().getModel();

            return new Promise((resolve, reject) => {
                oModel.read(sEntitySet, {
                    filters: [
                        new Filter("Bukrs", FilterOperator.EQ, sBukrs),
                        new Filter("Gjahr", FilterOperator.EQ, String(oPeriod.year)),
                        new Filter("Weeks", FilterOperator.EQ, this._padWeek(oPeriod.week)),
                        new Filter("Waers", FilterOperator.EQ, "KRW")
                    ],
                    success: (oData) => resolve(oData.results || []),
                    error: reject
                });
            });
        },

        _buildStatementItems(oPeriodData) {
            const oAmounts = PL_TYPES.reduce((oResult, oType) => {
                oResult[oType.key] = this._createAmountSet(oPeriodData, oType.key);
                return oResult;
            }, {});

            const oCorporateTax = this._createZeroAmountSet();
            const oGrossProfit = this._addAmountSets(oAmounts.SA, oAmounts.CO);
            const oOperatingProfit = this._addAmountSets(oGrossProfit, oAmounts.SG);
            const oPreTaxProfit = this._addAmountSets(
                this._addAmountSets(oOperatingProfit, oAmounts.OI),
                oAmounts.OE
            );
            const oNetProfit = this._addAmountSets(oPreTaxProfit, oCorporateTax);

            const aRows = [
                this._createStatementRow("SA", "매출액", "normal", oAmounts.SA, this._createDetailRows(oPeriodData, "SA"), true),
                this._createStatementRow("CO", "매출원가", "deduction", oAmounts.CO, this._createDetailRows(oPeriodData, "CO"), true),
                this._createStatementRow("GROSS_PROFIT", "매출총이익", "subtotal", oGrossProfit, [], false, true),
                this._createStatementRow("SG", "판매관리비", "deduction", oAmounts.SG, this._createDetailRows(oPeriodData, "SG"), true),
                this._createStatementRow("OPERATING_PROFIT", "영업이익(손실)", "subtotal", oOperatingProfit, [], false, true),
                this._createStatementRow("OI", "영업외수익", "addition", oAmounts.OI, this._createDetailRows(oPeriodData, "OI"), true),
                this._createStatementRow("OE", "영업외비용", "deduction", oAmounts.OE, this._createDetailRows(oPeriodData, "OE"), true),
                this._createStatementRow("PRE_TAX_PROFIT", "법인세차감전이익(손실)", "subtotal", oPreTaxProfit, [], false, true),
                this._createStatementRow("CORPORATE_TAX", "법인세비용", "deduction", oCorporateTax, [], false),
                this._createStatementRow("NET_PROFIT", "당기순이익(손실)", "total", oNetProfit, [], false, true)
            ];

            return aRows.filter((oRow) => oRow.forceShow || oRow.alwaysShow || !this._isZeroRow(oRow) || (oRow.children || []).length > 0);
        },

        _createStatementRow(sKey, sName, sRowType, oAmounts, aChildren = [], bForceShow = false, bAlwaysShow = false) {
            const aFilteredChildren = aChildren.filter((oChild) => !this._isZeroRow(oChild));

            // If this row should be forced to show (e.g. summary from ZCDS_D4_FI_0002)
            // but has no real children, add a hidden placeholder child so the Tree
            // will render an expander '>' even when amounts are zero.
            let aChildrenForModel = aFilteredChildren;

            if (bForceShow && aFilteredChildren.length === 0) {
                aChildrenForModel = [{
                    PL_type: sKey + "_PLACEHOLDER",
                    name: "",
                    rowType: "detail",
                    Waers: "KRW",
                    previous2Amount: 0,
                    previousAmount: 0,
                    currentAmount: 0,
                    isPlaceholder: true,
                    children: []
                }];
            }

            return Object.assign({
                PL_type: sKey,
                name: sName,
                rowType: sRowType,
                Waers: "KRW",
                forceShow: bForceShow,
                hasExpander: bForceShow || aFilteredChildren.length > 0,
                alwaysShow: bAlwaysShow,
                expanded: true,
                children: aChildrenForModel
            }, oAmounts);
        },

        _createDetailRows(oPeriodData, sPlType) {
            const oRowsByAccount = {};

            this._getPeriodKeys().forEach((sPeriodKey) => {
                const aRows = (oPeriodData[sPeriodKey] && oPeriodData[sPeriodKey].detail) || [];

                aRows
                    .filter((oRow) => oRow.PL_type === sPlType)
                    .forEach((oRow) => {
                        const sAccountKey = oRow.Saknr || oRow.Gltxt || sPlType;

                        if (!oRowsByAccount[sAccountKey]) {
                            oRowsByAccount[sAccountKey] = {
                                PL_type: sPlType,
                                Saknr: oRow.Saknr,
                                name: this._getDetailName(oRow),
                                rowType: "detail",
                                Waers: oRow.Waers || "KRW",
                                previous2Amount: 0,
                                previousAmount: 0,
                                currentAmount: 0,
                                children: []
                            };
                        }

                        oRowsByAccount[sAccountKey][sPeriodKey + "Amount"] += this._getDetailAmount(oRow);
                    });
            });

            return Object.keys(oRowsByAccount)
                .sort()
                .map((sAccountKey) => oRowsByAccount[sAccountKey])
                .filter((oRow) => !this._isZeroRow(oRow));
        },

        _getDetailName(oRow) {
            return String(oRow.Gltxt || "").trim();
        },

        _getDetailAmount(oRow) {
            return Number(oRow.h_amount || 0) - Number(oRow.s_amount || 0);
        },

        _createAmountSet(oPeriodData, sPlType) {
            return {
                previous2Amount: this._getAmount(oPeriodData.previous2 && oPeriodData.previous2.summary, sPlType),
                previousAmount: this._getAmount(oPeriodData.previous && oPeriodData.previous.summary, sPlType),
                currentAmount: this._getAmount(oPeriodData.current && oPeriodData.current.summary, sPlType)
            };
        },

        _createZeroAmountSet() {
            return {
                previous2Amount: 0,
                previousAmount: 0,
                currentAmount: 0
            };
        },

        _addAmountSets(oLeft, oRight) {
            return AMOUNT_FIELDS.reduce((oResult, sField) => {
                oResult[sField] = Number(oLeft[sField] || 0) + Number(oRight[sField] || 0);
                return oResult;
            }, {});
        },

        _isZeroRow(oRow) {
            return AMOUNT_FIELDS.every((sField) => Number(oRow[sField] || 0) === 0);
        },

        _getAmount(aRows, sPlType) {
            const oRow = (aRows || []).find((oItem) => oItem.PL_type === sPlType);
            return oRow ? Number(oRow.amount || 0) : 0;
        },

        _createPeriodMap(aPeriods) {
            return {
                previous2: aPeriods[0],
                previous: aPeriods[1],
                current: aPeriods[2]
            };
        },

        _getPeriodKeys() {
            return ["previous2", "previous", "current"];
        },

        _getDisplayPeriods(iYear, iWeek) {
            return [
                Object.assign({ key: "previous2" }, this._shiftFiscalWeek(iYear, iWeek, -2)),
                Object.assign({ key: "previous" }, this._shiftFiscalWeek(iYear, iWeek, -1)),
                Object.assign({ key: "current" }, this._shiftFiscalWeek(iYear, iWeek, 0))
            ];
        },

        _getRelativeFiscalWeek(oDate, iWeekOffset) {
            const oWeek = this._getIsoWeek(oDate);
            return this._shiftFiscalWeek(oWeek.year, oWeek.week, iWeekOffset);
        },

        _shiftFiscalWeek(iYear, iWeek, iOffset) {
            let iTargetYear = iYear;
            let iTargetWeek = iWeek + iOffset;

            while (iTargetWeek < 1) {
                iTargetYear -= 1;
                iTargetWeek += this._getWeeksInYear(iTargetYear);
            }

            while (iTargetWeek > this._getWeeksInYear(iTargetYear)) {
                iTargetWeek -= this._getWeeksInYear(iTargetYear);
                iTargetYear += 1;
            }

            const oEndDate = this._getIsoWeekEndDate(iTargetYear, iTargetWeek);
            const oStartDate = new Date(oEndDate.getTime());
            oStartDate.setUTCDate(oStartDate.getUTCDate() - 6);

            return {
                year: iTargetYear,
                week: iTargetWeek,
                weekText: this._padWeek(iTargetWeek),
                title: "제 " + iTargetWeek + "기",
                periodText: "제 " + iTargetWeek + "기 " + this._formatKoreanDate(oStartDate) + " ~ " + this._formatKoreanDate(oEndDate)
            };
        },

        _getIsoWeek(oDate) {
            const oUtcDate = new Date(Date.UTC(oDate.getFullYear(), oDate.getMonth(), oDate.getDate()));
            const iDay = oUtcDate.getUTCDay() || 7;
            oUtcDate.setUTCDate(oUtcDate.getUTCDate() + 4 - iDay);

            const iYear = oUtcDate.getUTCFullYear();
            const oYearStart = new Date(Date.UTC(iYear, 0, 1));
            const iWeek = Math.ceil((((oUtcDate - oYearStart) / 86400000) + 1) / 7);

            return {
                year: iYear,
                week: iWeek
            };
        },

        _getIsoWeekEndDate(iYear, iWeek) {
            const oJan4 = new Date(Date.UTC(iYear, 0, 4));
            const iJan4Day = oJan4.getUTCDay() || 7;
            const oWeek1Monday = new Date(oJan4.getTime());

            oWeek1Monday.setUTCDate(oJan4.getUTCDate() - iJan4Day + 1);

            const oWeekEndDate = new Date(oWeek1Monday.getTime());
            oWeekEndDate.setUTCDate(oWeek1Monday.getUTCDate() + ((iWeek - 1) * 7) + 6);

            return oWeekEndDate;
        },

        _formatKoreanDate(oDate) {
            return [
                oDate.getUTCFullYear() + "년",
                this._pad2(oDate.getUTCMonth() + 1) + "월",
                this._pad2(oDate.getUTCDate()) + "일"
            ].join(" ");
        },

        _getWeeksInYear(iYear) {
            return this._getIsoWeek(new Date(Date.UTC(iYear, 11, 28))).week;
        },

        _padWeek(vWeek) {
            const iWeek = Number(vWeek);
            return Number.isFinite(iWeek) ? String(iWeek).padStart(2, "0") : "";
        },

        _pad2(iValue) {
            return String(iValue).padStart(2, "0");
        }
    });
});
