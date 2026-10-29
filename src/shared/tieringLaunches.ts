import type { HistoricalLaunchMetadata, TieringFeature } from "./types.js";

export interface HistoricalLaunchRecord extends HistoricalLaunchMetadata {
  featureId: string;
}

const SOURCES = {
  bandHistory: "https://consumer.huawei.com/ph/stories/wearables/news-band/",
  sleep: "https://consumer.huawei.com/cn/support/content/zh-cn15963854/",
  sustainability2122: "https://consumer.huawei.com/content/dam/huawei-cbg-site/cn/mkt/campaign/sustainability/%E5%8D%8E%E4%B8%BA%E7%BB%88%E7%AB%AF%E5%8F%AF%E6%8C%81%E7%BB%AD%E5%8F%91%E5%B1%95%E6%8A%A5%E5%91%8A%EF%BC%882021-2022%EF%BC%89.pdf",
  sustainability2324: "https://consumer.huawei.com/content/dam/huawei-cbg-site/greate-china/cn/mkt/campaign/sustainability-new/231021/%E5%8D%8E%E4%B8%BA%E7%BB%88%E7%AB%AF%E5%8F%AF%E6%8C%81%E7%BB%AD%E5%8F%91%E5%B1%95%E6%8A%A5%E5%91%8A%EF%BC%882023-2024%EF%BC%89.pdf",
  ultimate: "https://consumer.huawei.com/cn/support/content/zh-cn15943099/",
  watch4: "https://consumer.huawei.com/cn/support/content/zh-cn15953902/",
  gt4: "https://consumer.huawei.com/cn/support/content/zh-cn15971720/",
  osa: "https://consumer.huawei.com/cn/support/content/zh-cn15990203/",
  gt5: "https://consumer.huawei.com/cn/support/content/zh-cn16015388/",
  d2: "https://consumer.huawei.com/cn/wearables/watch-d2/",
  watch5: "https://consumer.huawei.com/cn/wearables/watch-5/",
} as const;

type Seed = Omit<HistoricalLaunchMetadata, "suggestedTargetMonth"> & { suggestedTargetMonth: string };

function month(targetMonth: string, evidenceUrls: string[], note: string, confidence: "A" | "B" = "A"): Seed {
  return { suggestedTargetMonth: targetMonth, precision: "month", verificationStatus: "confirmed", confidence, evidenceUrls, note };
}

function halfYear(targetMonth: string, evidenceUrls: string[], note: string, confidence: "B" | "C" = "B"): Seed {
  return { suggestedTargetMonth: targetMonth, precision: "half-year", verificationStatus: "pending", confidence, evidenceUrls, note };
}

function year(targetMonth: string, evidenceUrls: string[], note: string): Seed {
  return { suggestedTargetMonth: targetMonth, precision: "year", verificationStatus: "pending", confidence: "C", evidenceUrls, note };
}

const rules: Array<{ from: number; to: number; seed: Seed }> = [
  { from: 1, to: 1, seed: year("2015-09", [SOURCES.sustainability2122], "单次腕上心率在早期智能穿戴代际已出现；中国大陆具体首发半年和月份仍待核验。") },
  { from: 2, to: 7, seed: halfYear("2018-09", [SOURCES.sustainability2122], "2018 年 HUAWEI WATCH GT 代际已具备 24 小时连续心率和基础专业运动能力；月份暂按 H2 规则存为 9 月。") },
  { from: 8, to: 8, seed: halfYear("2019-09", [SOURCES.sustainability2122], "TruSeen 3.5 在 2019 年引入单次血氧能力；月份暂按 H2 规则存为 9 月。") },
  { from: 9, to: 11, seed: halfYear("2020-09", [SOURCES.sustainability2122], "连续/自动血氧及趋势能力按 2020 年穿戴代际归档；中国大陆具体月份待核验。") },
  { from: 12, to: 12, seed: year("2014-04", [SOURCES.bandHistory], "TalkBand B1 代际已有基础睡眠识别；年份可信，半年和月份待中国大陆资料核验。") },
  { from: 13, to: 16, seed: year("2017-09", [SOURCES.sleep], "华为官方将 TruSleep 1.0 起点标为 2017 年；半年和月份仍待核验。") },
  { from: 17, to: 18, seed: month("2023-12", [SOURCES.osa], "睡眠呼吸暂停筛查软件中国说明书于 2023 年 12 月发布，作为公开可用时间锚点。") },
  { from: 19, to: 21, seed: halfYear("2019-09", [SOURCES.sustainability2122], "压力检测与趋势按 WATCH GT 2 代际归档；中国大陆具体月份待核验。") },
  { from: 22, to: 22, seed: month("2024-09", [SOURCES.gt5], "GT 5 中国首发代际公开三状态情绪健康能力。") },
  { from: 23, to: 25, seed: halfYear("2025-04", [SOURCES.watch5], "WATCH 5 中国首发代际公开 12 种情绪及相关趋势提醒；月份暂按 H1 规则存为 4 月。") },
  { from: 26, to: 28, seed: halfYear("2021-04", [SOURCES.sustainability2122], "WATCH 3 代际引入腕部皮肤温度相关能力；月份暂按 H1 规则存为 4 月。") },
  { from: 29, to: 31, seed: month("2018-09", [SOURCES.sustainability2122], "心脏健康研究于 2018 年 9 月启动，覆盖房颤/早搏研究型筛查。") },
  { from: 32, to: 34, seed: halfYear("2020-09", [SOURCES.sustainability2122], "TruSeen 4.0/医疗级 ECG 初代在 2020 年穿戴代际出现；月份暂按 H2 规则存为 9 月。") },
  { from: 35, to: 36, seed: month("2021-12", [SOURCES.sustainability2122], "HUAWEI WATCH D 于 2021 年 12 月在中国发布，公开腕部血压能力。") },
  { from: 37, to: 37, seed: month("2024-11", [SOURCES.d2], "WATCH D2 中国首发代际公开 24 小时动态血压能力。") },
  { from: 38, to: 38, seed: month("2021-12", [SOURCES.sustainability2122], "血压趋势与历史记录随 WATCH D 中国首发代际公开。") },
  { from: 39, to: 50, seed: halfYear("2017-09", [SOURCES.sustainability2122], "华为体脂秤早期代际已提供体重和基础身体成分；具体功能逐项首发月份待核验。", "C") },
  { from: 51, to: 53, seed: halfYear("2021-09", [SOURCES.sustainability2122], "节段身体成分及附加指标按体脂秤 Pro 代际归档；具体月份待核验。", "C") },
  { from: 54, to: 54, seed: month("2018-09", [SOURCES.sustainability2122], "心脏健康研究于 2018 年 9 月启动。") },
  { from: 55, to: 55, seed: halfYear("2020-09", [SOURCES.sustainability2122], "术后健康管理研究按 2020 年研究项目批次归档；月份待核验。", "C") },
  { from: 56, to: 56, seed: month("2023-05", [SOURCES.watch4], "血管健康风险评估随 WATCH 4 微体检代际在中国公开。", "B") },
  { from: 57, to: 57, seed: month("2021-12", [SOURCES.sustainability2122], "血压健康研究随 WATCH D 中国首发节点归档。", "B") },
  { from: 58, to: 58, seed: month("2023-05", [SOURCES.watch4], "高血压风险评估随 WATCH 4 微体检代际在中国公开。", "B") },
  { from: 59, to: 59, seed: month("2024-04", [SOURCES.sustainability2324], "官方可持续发展报告确认心血管健康趋势研究于 2024 年 4 月启动。") },
  { from: 60, to: 60, seed: halfYear("2024-09", [SOURCES.d2], "运动与血压研究按 WATCH D2 代际归档；中国大陆具体月份待进一步核验。", "C") },
  { from: 61, to: 61, seed: month("2023-05", [SOURCES.watch4], "高血糖风险评估随 WATCH 4 微体检代际在中国公开。") },
  { from: 62, to: 62, seed: halfYear("2021-09", [SOURCES.sustainability2122], "肝脏脂肪研究按体脂秤 Pro/创新研究代际归档；月份待核验。", "C") },
  { from: 63, to: 63, seed: month("2023-05", [SOURCES.watch4], "呼吸风险评估随 WATCH 4 微体检代际公开。", "B") },
  { from: 64, to: 64, seed: halfYear("2019-09", [SOURCES.sustainability2122], "TruSeen 3.5 在 2019 年引入高原低氧提醒；月份暂按 H2 规则存为 9 月。") },
  { from: 65, to: 69, seed: halfYear("2021-09", [SOURCES.sustainability2122], "相关创新研究按 2021 年公开研究批次归档；逐项中国大陆月份待核验。", "C") },
  { from: 70, to: 70, seed: month("2024-09", [SOURCES.gt5], "情绪状态识别研究随 GT 5 情绪健康能力在中国公开。", "B") },
  { from: 71, to: 71, seed: month("2024-02", [SOURCES.sustainability2324], "官方报告确认情绪与睡眠改善计划于 2024 年 2 月启动。") },
  { from: 72, to: 73, seed: halfYear("2021-09", [SOURCES.sustainability2122], "运动负荷和跑步生理研究按 GT 3/GT Runner 代际归档；月份待核验。") },
  { from: 74, to: 74, seed: month("2023-09", [SOURCES.gt4], "骑行模拟功率按 GT 4 中国首发运动能力归档。", "B") },
  { from: 75, to: 76, seed: month("2024-02", [SOURCES.sustainability2324], "官方报告确认卵巢健康相关研究于 2024 年 2 月启动。") },
  { from: 77, to: 77, seed: month("2023-09", [SOURCES.gt4], "女性生理周期 3.0 随 GT 4 中国首发代际公开。", "B") },
  { from: 78, to: 80, seed: halfYear("2025-04", [SOURCES.watch5], "健康摘要随 WATCH 5 中国首发代际公开；月份暂按 H1 规则存为 4 月。") },
  { from: 81, to: 82, seed: month("2023-05", [SOURCES.watch4], "60 秒微体检及报告解读随 WATCH 4 中国首发代际公开。") },
  { from: 83, to: 83, seed: month("2024-11", [SOURCES.d2], "动态血压报告随 WATCH D2 中国首发代际公开。") },
  { from: 84, to: 84, seed: year("2017-09", [SOURCES.sleep], "睡眠结构、评分与建议报告源于 2017 年 TruSleep 1.0；半年和月份待核验。") },
  { from: 85, to: 85, seed: halfYear("2017-09", [SOURCES.sustainability2122], "身体成分趋势报告按早期华为体脂秤代际归档；月份待核验。", "C") },
  { from: 86, to: 86, seed: halfYear("2020-09", [SOURCES.sustainability2122], "心电记录与结果报告随 ECG 初代公开；月份待核验。") },
  { from: 87, to: 88, seed: month("2021-12", [SOURCES.sustainability2122], "血压管理能力随 WATCH D 中国首发节点归档。", "B") },
  { from: 89, to: 90, seed: month("2023-09", [SOURCES.gt4], "科学减脂和体重趋势管理随 GT 4 中国首发代际公开。", "B") },
  { from: 91, to: 91, seed: month("2023-05", [SOURCES.watch4], "高血糖风险评估随 WATCH 4 微体检代际公开。") },
  { from: 92, to: 92, seed: month("2018-09", [SOURCES.sustainability2122], "房颤与早搏持续风险管理源于 2018 年 9 月心脏健康研究。") },
  { from: 93, to: 93, seed: month("2023-12", [SOURCES.osa], "睡眠呼吸暂停风险筛查软件中国说明书于 2023 年 12 月发布。") },
  { from: 94, to: 95, seed: month("2023-05", [SOURCES.watch4], "血管和呼吸风险评估随 WATCH 4 微体检代际公开。", "B") },
  { from: 96, to: 100, seed: halfYear("2018-09", [SOURCES.sustainability2122], "运动健康女性周期记录和预测按早期应用代际归档；中国大陆具体月份待核验。", "C") },
  { from: 101, to: 102, seed: month("2023-09", [SOURCES.gt4], "结合穿戴数据的生理周期 3.0 与提醒随 GT 4 中国首发代际公开。") },
  { from: 103, to: 105, seed: halfYear("2019-09", [SOURCES.sustainability2122], "呼吸训练和压力舒缓按 GT 2 代际归档；月份待核验。", "C") },
  { from: 106, to: 106, seed: month("2024-09", [SOURCES.gt5], "情绪提醒后的改善建议随 GT 5 情绪健康能力公开。", "B") },
  { from: 107, to: 107, seed: year("2017-09", [SOURCES.sleep], "睡眠改善建议源于 2017 年 TruSleep 1.0；半年和月份待核验。") },
  { from: 108, to: 108, seed: halfYear("2019-09", [SOURCES.sustainability2122], "睡眠模式联动提醒抑制按 GT 2 代际归档；月份待核验。", "C") },
  { from: 109, to: 114, seed: halfYear("2021-09", [SOURCES.sustainability2122], "家庭空间、共享和紧急联系人按 2021 年运动健康/穿戴代际归档；月份待核验。", "C") },
  { from: 115, to: 116, seed: halfYear("2021-04", [SOURCES.sustainability2122], "跌倒检测和 SOS 随 WATCH 3 中国首发代际公开；月份暂按 H1 规则存为 4 月。") },
  { from: 117, to: 121, seed: year("2015-09", [SOURCES.sustainability2122], "基础锻炼记录按早期 HUAWEI WATCH 代际归档；半年和月份待核验。") },
  { from: 122, to: 124, seed: halfYear("2018-09", [SOURCES.sustainability2122], "独立 GNSS、轨迹和路线能力按 WATCH GT 代际归档；月份待核验。") },
  { from: 125, to: 126, seed: halfYear("2020-09", [SOURCES.sustainability2122], "路线导入与返航按 GT 2 Pro 代际归档；月份待核验。", "C") },
  { from: 127, to: 129, seed: halfYear("2021-09", [SOURCES.sustainability2122], "跑步和健身课程按 GT 3/GT Runner 代际归档；月份待核验。") },
  { from: 130, to: 130, seed: halfYear("2018-09", [SOURCES.sustainability2122], "最大摄氧量按 WATCH GT 基础专业运动代际归档；月份待核验。", "C") },
  { from: 131, to: 138, seed: halfYear("2021-09", [SOURCES.sustainability2122], "训练负荷、恢复、跑力指数和跑姿等按 GT 3/GT Runner 代际归档；月份待核验。") },
  { from: 139, to: 139, seed: halfYear("2020-09", [SOURCES.sustainability2122], "高尔夫练习模式按 GT 2 Pro 代际归档；月份待核验。", "C") },
  { from: 140, to: 141, seed: month("2023-03", [SOURCES.ultimate], "高尔夫球场和智能辅助能力按 WATCH Ultimate 中国首发节点归档。", "B") },
  { from: 142, to: 144, seed: month("2023-03", [SOURCES.ultimate], "高阶潜水和户外探险能力随 WATCH Ultimate 于 2023 年 3 月在中国发布。") },
  { from: 145, to: 145, seed: month("2023-09", [SOURCES.gt4], "骑行功率能力按 GT 4 中国首发运动能力归档。", "B") },
  { from: 146, to: 148, seed: year("2014-04", [SOURCES.bandHistory], "TalkBand B1 代际已有步数、距离和卡路里等基础活动能力；年份可信，半年和月份待核验。") },
  { from: 149, to: 152, seed: month("2023-05", [SOURCES.watch4], "新活力三环及目标设置随 WATCH 4 中国首发代际公开。") },
  { from: 153, to: 153, seed: year("2014-04", [SOURCES.bandHistory], "TalkBand B1 代际已有久坐提醒；年份可信，半年和月份待核验。") },
  { from: 154, to: 156, seed: month("2023-05", [SOURCES.watch4], "新活力三环的提醒、统计和建议随 WATCH 4 中国首发代际公开。", "B") },
];

function featureNumber(featureId: string): number {
  const match = /^F(\d{3})$/.exec(featureId);
  if (!match) throw new Error(`无效分层功能 ID：${featureId}`);
  return Number(match[1]);
}

export function historicalLaunchForFeatureId(featureId: string): HistoricalLaunchRecord {
  const number = featureNumber(featureId);
  const rule = rules.find((item) => number >= item.from && number <= item.to);
  if (!rule) throw new Error(`分层功能 ${featureId} 缺少历史上线时间台账`);
  return { featureId, ...rule.seed, evidenceUrls: [...rule.seed.evidenceUrls] };
}

export function buildHistoricalLaunchLedger(features: readonly Pick<TieringFeature, "id">[]): HistoricalLaunchRecord[] {
  return features.map((feature) => historicalLaunchForFeatureId(feature.id));
}

export function normalizeHistoricalLaunchMetadata(raw: unknown): HistoricalLaunchMetadata | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const value = raw as Partial<HistoricalLaunchMetadata>;
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(value.suggestedTargetMonth ?? "")) return undefined;
  if (!["month", "half-year", "year"].includes(value.precision ?? "")) return undefined;
  if (!["confirmed", "pending"].includes(value.verificationStatus ?? "")) return undefined;
  if (!["A", "B", "C"].includes(value.confidence ?? "")) return undefined;
  return {
    suggestedTargetMonth: value.suggestedTargetMonth!,
    precision: value.precision!,
    verificationStatus: value.verificationStatus!,
    confidence: value.confidence!,
    evidenceUrls: [...new Set((value.evidenceUrls ?? []).filter((item): item is string => typeof item === "string" && /^https?:\/\//.test(item)))],
    note: typeof value.note === "string" ? value.note : "",
  };
}

export function formatHistoricalLaunch(metadata: HistoricalLaunchMetadata, targetMonth = metadata.suggestedTargetMonth): string {
  const [year, monthRaw] = targetMonth.split("-");
  const month = Number(monthRaw);
  if (metadata.precision === "month") return `${year}年${month}月`;
  if (metadata.precision === "half-year") return `${year}H${month <= 6 ? 1 : 2} · 月份待核验`;
  return `${year}年 · 半年及月份待核验`;
}

export function historicalLaunchStatusLabel(metadata?: HistoricalLaunchMetadata): string {
  if (!metadata) return "";
  if (metadata.precision === "month") return "月份已确认";
  return metadata.precision === "half-year" ? "月份待核验" : "半年及月份待核验";
}
