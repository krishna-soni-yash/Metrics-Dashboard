import * as React from 'react';
import { useState } from 'react';
import { Chart } from 'react-google-charts';
import {
    Stack,
    Text,
    Dropdown,
    IDropdownOption,
    PrimaryButton,
    DefaultButton,

    Dialog,
    DialogType,
    DialogFooter,

    Pivot,
    PivotItem,
    DetailsList,
    SelectionMode,
    CheckboxVisibility,
    IColumn,
    KeyCodes
} from '@fluentui/react';
import IGenericService from '../../../../services/IGenericServices';
import IProjectMetricsRepository from '../../../../repositories/repositoryInterface/IProjectMetricsRepository';
import { MetricsRepository } from '../../../../repositories/MetricsRepository';
import { getMetricsFromProjectMetrics } from '../../../../repositories/MetricsRepository';
import { GenericService } from '../../../../services/GenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import IWorklogManagmentRepository from '../../../../repositories/repositoryInterface/IWorklogManagmentRepository';
import { WorklogManagmentRepository } from '../../../../repositories/WorklogManagmentRepository';
import { getWorkLogManagementValues } from '../../../../repositories/WorklogManagmentRepository';
import { getManagementEffortLogValues, ManagemnetEffortLogRepository } from '../../../../repositories/ManagemnetEffortLogRepository';
import { getTaskManagementLogValues, TaskManagementLogRepository } from '../../../../repositories/TaskManagementLogRepository';
import IManagemnetEffortLogRepository from '../../../../repositories/repositoryInterface/IManagemnetEffortLogRepository';
import ITaskManagementLogRepository from '../../../../repositories/repositoryInterface/ITaskManagementLogRepository';
import ITaskManagementRepository from '../../../../repositories/repositoryInterface/ITaskManagementRepository';
import { getTaskManagementValues, TaskManagementRepository } from '../../../../repositories/TaskManagementRepository';
import { CSIRepository, getCSIValues } from '../../../../repositories/CSIRepository';
import ICSIRepository from '../../../../repositories/repositoryInterface/ICSIRespository';
import { RAIDLogRepository, getRiskValues } from '../../../../repositories/RAIDLogRepository';
import IRAIDLogRepository from '../../../../repositories/repositoryInterface/IRAIDLogRepository';
import ITestingDefectsRepository from '../../../../repositories/repositoryInterface/ITestingDefectsRepository';
import { ITestingDefects } from '../../../../Models/ITestingDefects';
import { getTestingDefectsValues, TestingDefectsRepository } from '../../../../repositories/TestingDefectsRepository';
import ICodeReviewDefectsRepository from '../../../../repositories/repositoryInterface/ICodeReviewDefectsRepository';
import { CodeReviewDefectRepository, getCodeReviewDefectsValues } from '../../../../repositories/CodeReviewDefects';
import IPPOApproversRepository from '../../../../repositories/repositoryInterface/IPPOApproversRepository';
import { PPOApproversRepository, getAllItems as getPPOApproversValues } from '../../../../repositories/PPOApproversRepositroy';
import { IAMSEffortLog } from '../../../../Models/IAMSEffortLog';
import { getIAMSEffortLogValues } from '../../../../repositories/AMSRepository';
import IAMSEffortLogRepository from '../../../../repositories/repositoryInterface/IAMSEffortLogRespository';
import { AMSEffortLogRepository } from '../../../../repositories/AMSRepository';
import { set } from '@microsoft/sp-lodash-subset/lib/index';
import IFacilitationReportRepository from '../../../../repositories/repositoryInterface/IFacilitationReportInterface';
import { FacilitationReportRepository, getFacilitationValues } from '../../../../repositories/FacilitationReportRepository';
import IRCARepository from '../../../../repositories/repositoryInterface/IRCARepository';
import { getRCAItems } from '../../../../repositories/RCARepository';
import { RCARepository } from '../../../../repositories/RCARepository';
//import { IMetrics } from '../../../../Models/IMetrics';

interface DashboardProps {
    //onSubmit?: (data: any) => void;
    //initialData?: any;
    context?: WebPartContext;
}
export interface MonthlyData {
    ActualEffort: number;
    PlannedEffort: number;
    size: number;
    plannedDuration: number;
    codeReviewDefects: number;

}

const projectTypeLookup: Record<string, string> = {
    devm: 'DevM',
    dev: 'Dev',
    agile: 'Agile',
    ams: 'AMS'
};

const normalizeProjectType = (projectType?: string | null): string | undefined => {
    const raw = (projectType ?? '').trim();
    if (!raw) {
        return undefined;
    }

    const canonical = projectTypeLookup[raw.toLowerCase()];
    if (canonical) {
        return canonical;
    }

    if (raw.toLowerCase() === 'all') {
        return undefined;
    }

    return raw;
};


export default function Dashboard({ context }: DashboardProps): JSX.Element {
    // Filters state
    const [projectTypeSelection, setProjectTypeSelection] = React.useState<{ key?: string; canonical?: string }>(() => ({
        key: 'devm',
        canonical: normalizeProjectType('devm')
    }));
    const [selectedProjectType, setSelectedProjectType] = React.useState('devm');
    const normalizedProjectType = React.useMemo(() => normalizeProjectType(selectedProjectType), [selectedProjectType]);
    const [selectedMonth, setselectedMonth] = useState<string | undefined>(new Date().getMonth().toString());
    //const [metric, setMetric] = useState<string | undefined>('Velocity');
    //const [startDate, setStartDate] = useState<Date | undefined>(undefined);
    //const [endDate, setEndDate] = useState<Date | undefined>(undefined);
    const [MetricsData, setMetricsData] = React.useState<any[]>([]);
    const [WorkLogData, setWorkLogData] = React.useState<any[]>([]);
    const [TaskManagementData, setTaskManagementData] = React.useState<any[]>([]);
    const [ManagementEffortLogData, setManagementEffortLogData] = React.useState<any[]>([]);
    const [ManagementTaskLogData, setManagementTaskLogData] = React.useState<any[]>([]);
    // NEW: track active pivot/tab
    const [selectedPivotKey, setSelectedPivotKey] = useState<string>('Customer');
    // NEW: visibility per chart and active details for selected chart
    const [visibleCharts, setVisibleCharts] = useState<Record<string, boolean>>({});
    // NEW: CSI state
    const [csiData, setCsiData] = React.useState<any[]>([]);
    const [RAIDdata, setRAIDdata] = React.useState<any[]>([]);
    const [riskChartData, setRiskChartData] = React.useState<any[]>([]);
    const [defectsData, setDefectsData] = React.useState<ITestingDefects[]>([]);
    const [CodeReviewDefectsData, setCodeReviewDefectsData] = React.useState<any[]>([]);
    // NEW: date validation error for CSI loaders
    //const [dateError, setDateError] = React.useState<string | null>(null);
    const [filteredWorkLogData, setFilteredWorkLogData] = React.useState<any[]>([]);
    //const [ScheduleVariation, setScheduleVariation] = React.useState<any[]>([]);
    const [CSITrendData, setCSITrendData] = React.useState<any[]>([]);

    const [MeanEffortVariation, setMeanEffortVariation] = React.useState<number>(0);
    const [MeanSV, setMeanSV] = React.useState<number>(0);
    const [MeanOverallProductivity, setMeanOverallProductivity] = React.useState<number>(0);
    const [StandardDeviationOfEV, setStandardDeviationOfEV] = React.useState<number>(0);
    const [StandardDeviationSV, setStandardDeviationOfSV] = React.useState<number>(0);
    const [EffortDistributionData, setEffortDistributionData] = React.useState<any[]>([]);
    const [efforVatiationChartData, setefforVatiationChartData] = React.useState<any[]>([]);
    const [scheduledVatiationChartData, setscheduledVatiationChartData] = React.useState<any[]>([]);
    const [OverAllProductivitySize, setOverAllProductivitySize] = React.useState<any[]>([]);
    const [ActualEffortOverAllProductivity, setActualEffortOverAllProductivity] = React.useState<any[]>([]);
    const [OverallProductivityChartData, setOverallProductivityChartData] = React.useState<any[]>([]);
    const [RequirementAnalysisEffortDensityChartData, setRequirementAnalysisEffortDensityChartData] = React.useState<any[]>([]);
    const [CodingProductivityChartData, setCodingProductivityChartData] = React.useState<any[]>([]);
    const [CodeReviewEffortDensityChartData, setCodeReviewEffortDensityChartData] = React.useState<any[]>([]);
    const [CodeReworkEffortDensityChartData, setCodeReworkEffortDensityChartData] = React.useState<any[]>([]);
    const [UnitTestingEffortDensityChartData, setUnitTestingEffortDensityChartData] = React.useState<any[]>([]);
    const [TestingExecutionEffortDensityChartData, setTestingExecutionEffortDensityChartData] = React.useState<any[]>([]);
    const [CodeReviewDefectDensityChartData, setCodeReviewDefectDensityChartData] = React.useState<any[]>([]);
    const [PostDeliveryDefectDensityChartData, setPostDeliveryDefectDensityChartData] = React.useState<any[]>([]);
    const [TestingPhaseDefectDensityChartData, setTestingPhaseDefectDensityChartData] = React.useState<any[]>([]);
    const [defectCount, setDefectCount] = React.useState<number>(0);
    const [codeReviewDefectCount, setCodeReviewDefectCount] = React.useState<number>(0);
    const [PPOApproversData, setPPOApproversData] = React.useState<any[]>([]);
    const [AMSEffortLogData, setAMSEffortLogData] = React.useState<IAMSEffortLog[]>([]);
    const [resourceUtilization, setResourceUtilization] = React.useState<number>(0);
    const [ResourceUtilizationTrendChartData, setResourceUtilizationTrendChartData] = React.useState<any[]>([]);
    const [CostOfQualityTrendChartData, setCostOfQualityTrendChartData] = React.useState<any[]>([]);
    const [costOfQualityTotal, setcostOfQualityTotal] = React.useState<number>(0);
    const [defectsDetailsListItems, setDefectsDetailsListItems] = React.useState<any[]>([]);
    const [defectsDetailsListColumns, setDefectsDetailsListColumns] = React.useState<IColumn[]>([]);
    const [facilitationReportData, setFacilitationReportData] = React.useState<any[]>([]);
    const [facilitationReportDataForChart, setFacilitationReportDataForChart] = React.useState<any[]>([]);
    const [agingFindingsData, setagingFindingsData] = React.useState<any[]>([]);
    const [pciChartData, setpciChartData] = React.useState<any[]>([]);
    const [RCAItems, setRCAItems] = React.useState<any[]>([]);
    // new KPI counts
    // const [openRootCauseCount, setOpenRootCauseCount] = React.useState<number>(0);
    // const [openIssuesCount, setOpenIssuesCount] = React.useState<number>(0);
    // const [openActionItemsCount, setOpenActionItemsCount] = React.useState<number>(0);

    // React.useEffect(() => {
    //     if (!Array.isArray(facilitationReportData)) {
    //         setOpenRootCauseCount(0);
    //         setOpenIssuesCount(0);
    //         setOpenActionItemsCount(0);
    //         return;
    //     }
    //     const norm = (s: any) => String(s ?? '').trim().toLowerCase();
    //     const isOpen = (it: any) => {
    //         const s = norm(it?.Status);
    //         return s !== 'closed' && s !== 'resolved' && s !== 'done';
    //     };
    //     setOpenRootCauseCount(facilitationReportData.filter(f => norm(f?.Category) === 'root cause' && isOpen(f)).length);
    //     setOpenIssuesCount(facilitationReportData.filter(f => norm(f?.Category) === 'issue' && isOpen(f)).length);
    //     setOpenActionItemsCount(facilitationReportData.filter(f => ['action item','action','actionitem'].indexOf(norm(f?.Category)) && isOpen(f)).length);
    // }, [facilitationReportData]);

    const latestProjectTypeRef = React.useRef<string | undefined>(selectedProjectType);

    React.useEffect(() => {
        latestProjectTypeRef.current = selectedProjectType;
        setProjectTypeSelection({ key: selectedProjectType, canonical: normalizedProjectType });
    }, [normalizedProjectType, selectedProjectType]);

    const handleProjectTypeChange = (option?: IDropdownOption) => {
        const key = option?.key as string | undefined;
        if (key) {
            setSelectedProjectType(key);
        } else {
            setSelectedProjectType('');
        }
    };

    // compute latest CSI value (from csiData) — pick common field names and latest by CSATAquiredDate (fallbacks)
    const latestCsiValue: { value: string; valueRaw: any }[] = (() => {
        // ...no CSI data
        if (!csiData || csiData.length === 0) return [];
        const CSIForMonth = csiData.filter((it: any) => {
            return new Date(it.CSATAquiredDate).getMonth().toString() === selectedMonth;
        });

        if (!CSIForMonth.length) {
            return [];
        }

        const latest = CSIForMonth
            .slice()
            .sort(
                (a, b) =>
                    new Date(a.CSATAquiredDate).getTime() -
                    new Date(b.CSATAquiredDate).getTime()
            )
        [CSIForMonth.length - 1];

        return ([{
            value: latest.Title +
                " <span style='color: rgb(102, 102, 102)!important;font-weight: normal !important;font-size: 12px !important'>as on " +
                new Date(latest.CSATAquiredDate).toLocaleDateString() +
                '</span>',
            valueRaw: latest.Title
        }]





        );


    })();


    //   const ratingData = [
    //     ['Month', 'Rating', 'Lower Limit', 'Upper Limit'],
    //     ['May', 4.6, 4.0, 5.0],
    //     ['Jun', 4.7, 4.0, 5.0],
    //     ['Jul', 4.8, 4.0, 5.0],
    //     ['Aug', 4.7, 4.0, 5.0],
    //     ['Sep', 4.9, 4.0, 5.0],
    // ];
    // build KPIs from MetricsData; override CSI value with latestCsiValue when available
    const desiredOrder = [
        'Customer Satisfaction Index',
        'Schedule Variation',
        'Effort Variation',
        'Post Delivery Defects',
        'Overall Productivity',
        'Internal Defects',
        'Cost of Quality',
        'Resource Utilization'
    ];

    let rawKpis: any[] = []
    if (MetricsData.length > 0) {
        rawKpis = (MetricsData || []).map(m => {

            try {
                const safe = (v: any) => (v === undefined || v === null) ? null : v;
                let goalLabel = m.goal;
                // if (safe(m?.LSL) != null && safe(m?.USL) != null) goalLabel = `${m.LSL} - ${m.USL}`;
                // else if (safe(m?.LSL) != null) goalLabel = `>= ${m.LSL}`;
                // else if (safe(m?.USL) != null) goalLabel = `<= ${m.USL}`;

                let value = String(m?.value ?? '');
                let status = 'red';
                const title = String(m?.title ?? m?.Metrics ?? m?.id ?? 'Unknown');
                const titleNormalized = title.trim().toLowerCase();


                if (titleNormalized === 'customer satisfaction index' && latestCsiValue.length > 0) {
                    value = String(latestCsiValue[0].value || '');
                    const raw = Number(latestCsiValue[0].valueRaw);
                    status = (raw >= (m?.LSL ?? -Infinity) && raw <= (m?.USL ?? Infinity)) ? 'green' : 'red';
                    return { id: m?.id, title, value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                }

                if (selectedProjectType?.toLocaleLowerCase() === 'devm') {

                    if (titleNormalized === 'effort variation_month' && MeanEffortVariation != null && StandardDeviationOfEV != null) {
                        value = `<span style='color:#666;font-size:12px'>Mean:</span>${MeanEffortVariation}<br/><span style='color:#666;font-size:12px'>Std Dev:</span>${StandardDeviationOfEV}`;
                        status = (MeanEffortVariation >= (m?.LSL ?? -Infinity) && MeanEffortVariation <= (m?.USL ?? Infinity)) ? 'green' : 'red';
                        return { id: m?.id, title, value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                    }


                    if (titleNormalized === 'schedule variation_month' && MeanSV != null && StandardDeviationSV != null) {
                        value = `<span style='color:#666;font-size:12px'>Mean:</span>${MeanSV}<br/><span style='color:#666;font-size:12px'>Std Dev:</span>${StandardDeviationSV}`;
                        status = (MeanSV >= (m?.LSL ?? -Infinity) && MeanSV <= (m?.USL ?? Infinity)) ? 'green' : 'red';
                        return { id: m?.id, title, value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                    }

                    if (titleNormalized === 'overall productivity_month' && MeanOverallProductivity != null && OverAllProductivitySize != null && ActualEffortOverAllProductivity != null) {
                        value = `<span style='color:#666;font-size:12px'>Effort:</span>${ActualEffortOverAllProductivity}<br/><span style='color:#666;font-size:12px'>Size:</span>${OverAllProductivitySize}`;
                        status = (MeanOverallProductivity >= (m?.LSL ?? -Infinity) && MeanOverallProductivity <= (m?.USL ?? Infinity)) ? 'green' : 'red';
                        return { id: m?.id, title, value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                    }


                    if (titleNormalized === 'delivered defect density_month' || titleNormalized === 'delivered defect density (post production)') {
                        const cnt = defectCount || 0;
                        value = `<span style='color:#666;font-size:12px'>Defects Count:</span>${cnt}`;
                        status = (cnt >= (m?.LSL ?? -Infinity) && cnt <= (m?.USL ?? Infinity)) ? 'green' : 'red';

                        return { id: m?.id, title: 'Post Delivery Defects', value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                    }


                    if (titleNormalized === 'defect density_month' || titleNormalized === 'defect_density' || titleNormalized === 'internal defect density') {
                        const cnt = codeReviewDefectCount || 0;
                        value = `<span style='color:#666;font-size:12px'>Defects Count:</span>${cnt}`;
                        status = (cnt >= (m?.LSL ?? -Infinity) && cnt <= (m?.USL ?? Infinity)) ? 'green' : 'red';
                        return { id: m?.id, title: 'Internal Defects', value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                    }
                }
                if (selectedProjectType?.toLocaleLowerCase() === 'dev') {
                    if (titleNormalized === 'effort variation' && MeanEffortVariation != null && StandardDeviationOfEV != null) {
                        value = `<span style='color:#666;font-size:12px'>Mean:</span>${MeanEffortVariation}<br/><span style='color:#666;font-size:12px'>Std Dev:</span>${StandardDeviationOfEV}`;
                        status = (MeanEffortVariation >= (m?.LSL ?? -Infinity) && MeanEffortVariation <= (m?.USL ?? Infinity)) ? 'green' : 'red';
                        return { id: m?.id, title, value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                    }


                    if (titleNormalized === 'schedule variation' && MeanSV != null && StandardDeviationSV != null) {
                        value = `<span style='color:#666;font-size:12px'>Mean:</span>${MeanSV}<br/><span style='color:#666;font-size:12px'>Std Dev:</span>${StandardDeviationSV}`;
                        status = (MeanSV >= (m?.LSL ?? -Infinity) && MeanSV <= (m?.USL ?? Infinity)) ? 'green' : 'red';
                        return { id: m?.id, title, value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                    }

                    if (titleNormalized === 'overall productivity' && MeanOverallProductivity != null && OverAllProductivitySize != null && ActualEffortOverAllProductivity != null) {
                        value = `<span style='color:#666;font-size:12px'>Effort:</span>${ActualEffortOverAllProductivity}<br/><span style='color:#666;font-size:12px'>Size:</span>${OverAllProductivitySize}`;
                        status = (MeanOverallProductivity >= (m?.LSL ?? -Infinity) && MeanOverallProductivity <= (m?.USL ?? Infinity)) ? 'green' : 'red';
                        return { id: m?.id, title, value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                    }


                    if (titleNormalized === 'delivered defect density' || titleNormalized === 'delivered defect density (post production)') {
                        const cnt = defectCount || 0;
                        value = `<span style='color:#666;font-size:12px'>Defects Count:</span>${cnt}`;
                        status = (cnt >= (m?.LSL ?? -Infinity) && cnt <= (m?.USL ?? Infinity)) ? 'green' : 'red';

                        return { id: m?.id, title: 'Post Delivery Defects', value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                    }


                    if (titleNormalized === 'defect density' || titleNormalized === 'defect_density' || titleNormalized === 'internal defect density') {
                        const cnt = codeReviewDefectCount || 0;
                        value = `<span style='color:#666;font-size:12px'>Defects Count:</span>${cnt}`;
                        status = (cnt >= (m?.LSL ?? -Infinity) && cnt <= (m?.USL ?? Infinity)) ? 'green' : 'red';
                        return { id: m?.id, title: 'Internal Defects', value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                    }
                }




                if (titleNormalized === 'resource utilization' || titleNormalized === 'resource_utilization') {

                    value = `${resourceUtilization.toFixed(2)}%`;
                    status = (resourceUtilization >= (m?.LSL ?? -Infinity) && resourceUtilization <= (m?.USL ?? Infinity)) ? 'green' : 'red';
                    return { id: m?.id, title: m?.title, value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                }
                if (titleNormalized === 'cost of quality' || titleNormalized === 'cost_of_quality') {

                    value = `${costOfQualityTotal.toFixed(2)}%`;
                    status = (costOfQualityTotal >= (m?.LSL ?? -Infinity) && costOfQualityTotal <= (m?.USL ?? Infinity)) ? 'green' : 'red';
                    return { id: m?.id, title: m?.title, value, goal: goalLabel, status, pivotKey: m?.pivotKey };
                }

                // default fallback KPI
                return {
                    id: m?.id,
                    title,
                    value,
                    goal: goalLabel,
                    status,
                    pivotKey: m?.pivotKey,
                };
            } catch (err) {
                // defensive fallback — don't break render if one metric fails
                console.error('KPI mapping error', err, m);
                return { id: m?.id ?? Math.random(), title: String(m?.title ?? m?.Metrics ?? 'Metric'), value: String(m?.value ?? ''), goal: '', status: 'red', pivotKey: m?.pivotKey ?? '' };
            }
        });
    } else {
        // No metrics available — provide placeholder KPI objects so UI and CSV export still work
        rawKpis = desiredOrder.map(name => ({
            id: name,
            title: name,
            value: '0',
            goal: '',
            status: 'red',
            pivotKey: ''
        }));
    }

    // Reorder KPIs to the requested display priority (case-insensitive).
    // Desired order: Customer Satisfaction Index, Schedule Variation, Effort Variation,
    // Post Delivery Defects, Overall Productivity, Internal Defects, Cost of Quality, Resource Utilization


    const remaining = [...rawKpis];
    const orderedKpis: any[] = [];
    desiredOrder.forEach(name => {
        let idx = -1;
        for (let i = 0; i < remaining.length; i++) {
            if (String(remaining[i].title || '').toLowerCase() === name.toLowerCase()) {
                idx = i;
                break;
            }
        }
        if (idx !== -1) {
            orderedKpis.push(remaining.splice(idx, 1)[0]);
        }
    });
    // Append any KPIs not listed in desiredOrder, preserving (or sorting) them as-is
    orderedKpis.push(...remaining);

    const kpis = orderedKpis;

    // Chart data (react-google-charts format)
    // velocity with overall productivity limits (lower=5, upper=25)
    // const velocityData = [
    //     ['Sprint', 'Velocity', 'Prod Lower', 'Prod Upper'],
    //     ['Sprint 1', 5, 5, 25],
    //     ['Sprint 2', 7, 5, 25],
    //     ['Sprint 3', 8, 5, 25],
    //     ['Sprint 4', 6, 5, 25],
    //     ['Sprint 5', 9, 5, 25],
    // ];

    // const resourceData = [
    //     ['Sprint', 'Velocity', 'Prod Lower', 'Prod Upper'],
    //     ['Sprint 1', 5, 5, 25],
    //     ['Sprint 2', 7, 5, 25],
    //     ['Sprint 3', 8, 5, 25],
    //     ['Sprint 4', 6, 5, 25],
    //     ['Sprint 5', 9, 5, 25],
    // ];

    // // defect density with lower/upper limits (lower=0, upper=2.0)
    // const defectData = [
    //     ['Sprint', 'Defect Density', 'Defect Lower', 'Defect Upper'],
    //     ['Sprint 1', 0.3, 0, 2.0],
    //     ['Sprint 2', 0.25, 0, 2.0],
    //     ['Sprint 3', 0.4, 0, 2.0],
    //     ['Sprint 4', 0.35, 0, 2.0],
    //     ['Sprint 5', 0.2, 0, 2.0],
    // ];

    // // resource/customer rating with lower/upper limits (lower=4.0, upper=5.0)


    // // weekly report with lower/upper limits (lower=3.5, upper=5.0)
    // const weekData = [
    //     ['Week', 'Weekly Report', 'Lower Limit', 'Upper Limit'],
    //     ['Week1', 4.9, 3.5, 5.0],
    //     ['Week2', 5, 3.5, 5.0],
    //     ['Week3', 4, 3.5, 5.0],
    //     ['Week4', 3.5, 3.5, 5.0],
    // ];

    // const riskData = [
    //     ['Week', 'Risks Open'],
    //     ['Week1', 14],
    //     ['Week2', 15],
    //     ['Week3', 40],
    //     ['Week4', 3],
    // ];

    // NEW: findings summary data provided by user
    // const findingsTypeData = [
    //     ['Type', 'Count'],
    //     ['NCs', 101],
    //     ['Observations', 356],
    //     ['Facilitation', 279],
    // ];

    // const contributorsData = [
    //     ['Entity', 'Count'],
    //     ['IT IS', 43],
    //     ['IT IS Europe', 35],
    //     ['FPA', 27],
    //     ['Dormer', 39],
    //     ['Americhem AI platform', 23],
    //     ['RLUS Analytics', 23],
    // ];

    // NEW: critical findings breakdown (from user)
    // const criticalBreakdownData = [
    //     ['Category', 'Count'],
    //     ['Technical', 54],
    //     ['Non-Technical', 104],
    // ];

    // NEW: data for Critical Risk chart (Critical first to emphasize)
    // const criticalRiskData = [
    //     ['Severity', 'Count'],
    //     ['Critical', 5],
    //     ['High', 12],
    //     ['Medium', 20],
    //     ['Low', 8],
    // ];
    // NEW: Aging findings data (example buckets summing to total 736)
    // const agingFindingsData = [
    //     ['Aging', 'Findings'],
    //     ['>=7 days', 10],
    //     ['<=30 days', 2],
    //     ['>30 days', 5],
    // ];

    // PCI (Process Compliance Index) data is computed from Management effort and facilitation findings
    //const pciChartData = buildPCIChartData();

    const pciOptions = {
        title: 'Process Compliance Index (PCI)',
        legend: { position: 'none' },
        seriesType: 'line',
        series: { 0: { color: '#1e88e5', lineWidth: 2, pointSize: 6 } },
        colors: ['#1e88e5'],
        annotations: {
            alwaysOutside: true,
            textStyle: { fontSize: 10, color: '#222', bold: false },
            stem: { color: 'transparent', length: 0 }
        },
        hAxis: { textStyle: { color: '#555' } },
        vAxis: { viewWindow: { min: -160, max: 40 }, gridlines: { color: '#eee' } },
        chartArea: { left: 60, top: 40, right: 20, bottom: 40 },
        dataOpacity: 0.95
    };

    // Simple CSV export (example for visible KPI + selected metric)
    const exportToCsv = () => {
        const rows = [
            ['Metric', 'Value'],
            ...kpis.map(k => [k.title, k.value]),
        ];
        const csv = rows.map(r => r.join(',')).join('\n');
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'metrics_export.csv';
        a.click();
        URL.revokeObjectURL(url);
    };

    // Fluent options for dropdowns
    const ProjectType: IDropdownOption[] = [
        { key: 'devm', text: 'DevM' },
        { key: 'dev', text: 'Dev' },
        { key: 'agile', text: 'Agile' },
        { key: 'ams', text: 'AMS' },
    ];
    const Months: IDropdownOption[] = [
        { key: '0', text: 'January' },
        { key: '1', text: 'February' },
        { key: '2', text: 'March' },
        { key: '3', text: 'April' },
        { key: '4', text: 'May' },
        { key: '5', text: 'June' },
        { key: '6', text: 'July' },
        { key: '7', text: 'August' },
        { key: '8', text: 'September' },
        { key: '9', text: 'October' },
        { key: '10', text: 'November' },
        { key: '11', text: 'December' },
    ];

    // const metricOptions: IDropdownOption[] = [
    //     { key: 'Velocity', text: 'Velocity' },
    //     { key: 'Defect', text: 'Defect Density' },
    //     { key: 'Resource', text: 'Resource Utilization' },
    //     { key: 'Weekly', text: 'Weekly Report' },
    //     { key: 'Risk', text: 'Risk' },
    // ];

    // Basic style objects (keeps layout similar to original)
    const containerStyle: React.CSSProperties = {
        fontFamily: "'Inter', Arial, sans-serif",
        margin: 0,
        backgroundColor: '#f3f6fb',
        color: '#222',
        padding: '24px',
        paddingBottom: 40,
        // width: '100vw',
        // minHeight: '100vh',
        boxSizing: 'border-box',
    };

    // centered content wrapper (bootstrap-like container)
    const contentStyle: React.CSSProperties = {
        //maxWidth: 1200,
        margin: '0 auto',
        padding: 20,
        boxSizing: 'border-box',
        background: '#ffffff',
        borderRadius: 12,
        boxShadow: '0 8px 30px rgba(15, 20, 30, 0.06)',
    };

    const headerStyle: React.CSSProperties = {
        background: 'transparent',
        padding: 8,
        textAlign: 'left',
        marginBottom: 6,
    };

    // updated KPI area style: single row, no wrap, allow horizontal scroll on small screens
    const kpiAreaStyle: React.CSSProperties = {
        display: 'flex',
        gap: 12,
        flexWrap: 'wrap',        // keep all cards on one line
        justifyContent: 'space-between',
        alignItems: 'stretch',
        //overflowX: 'auto',         // allow scroll on narrow screens
        paddingBottom: 8,
    };

    // updated KPI card to take 1/6 width
    const kpiCardStyle = (): React.CSSProperties => ({
        background: '#fff',
        padding: 14,
        borderRadius: 10,
        flex: '0 0 21%',       // fixed sixth-per-row
        minWidth: 180,             // keep readable on small screens
        textAlign: 'left',
        boxShadow: '0 2px 8px rgba(0,0,0,0.04)',
        color: '#333',
        display: 'flex',
        flexDirection: 'column',
        justifyContent: 'space-between',
        transition: 'transform 0.12s ease, box-shadow 0.12s ease',
    });

    // const valueStyle = (status?: string): React.CSSProperties => ({
    //     fontSize: 26,
    //     fontWeight: 700,
    //     marginTop: 8,
    //     marginBottom: 6,
    //     color: status === 'green' ? '#28a745' : '#dc3545',
    // });

    // UPDATED: remove width/maxWidth on container; use 50% per chart so two charts fill viewport exactly
    const chartsContainerStyle: React.CSSProperties = {
        display: 'flex',
        gap: 16,
        margin: '20px 0',
        width: '100%',
        flexWrap: 'wrap',
        alignItems: 'stretch',
        boxSizing: 'border-box',
        padding: 0,
    };

    const chartBoxStyle: React.CSSProperties = {
        background: '#fff',
        flex: '1 1 48%',
        minWidth: 280,
        minHeight: 260,
        borderRadius: 10,
        boxShadow: '0 2px 10px rgba(12, 20, 40, 0.04)',
        padding: 12,
        boxSizing: 'border-box',
        display: 'flex',
        flexDirection: 'column',
    };

    //   // NEW: full-width chart box for the Critical Risk chart
    //   const chartFullBoxStyle: React.CSSProperties = {
    //     background: '#fff',
    //     width: '100%',
    //     minHeight: 260,
    //     borderRadius: 10,
    //     boxShadow: '0 2px 8px rgba(0,0,0,0.05)',
    //     padding: 12,
    //     boxSizing: 'border-box',
    //     display: 'flex',
    //     flexDirection: 'column',
    //     justifyContent: 'center',
    //   };

    // NEW: two-column container and half-box style for side-by-side charts below nav
    const belowNavContainerStyle: React.CSSProperties = {
        display: 'flex',
        gap: 16,
        marginTop: 20,
        width: '100%',
        flexWrap: 'wrap',
        boxSizing: 'border-box',
    };

    const halfChartBoxStyle: React.CSSProperties = {
        background: '#fff',
        flex: '1 1 48%',
        minWidth: 300,
        borderRadius: 10,
        boxShadow: '0 2px 10px rgba(12, 20, 40, 0.04)',
        padding: 12,
        boxSizing: 'border-box',
        minHeight: 260,
        display: 'flex',
        flexDirection: 'column',
    };

    // NEW: aging chart full-width box (placed below the two-column findings section)
    const agingBoxStyle: React.CSSProperties = {
        background: '#fff',
        width: '100%',
        borderRadius: 10,
        boxShadow: '0 2px 10px rgba(12, 20, 40, 0.04)',
        padding: 12,
        boxSizing: 'border-box',
        marginTop: 16,
    };


    const [activeDetails, setActiveDetails] = useState<{
        key: string;
        title: string;
        columns: IColumn[];
        items: any[];
    } | null>(null);

    // Dialog state for KPI -> Pivot content
    const [isDialogOpen, setIsDialogOpen] = useState<boolean>(false);
    const [dialogPivotKey, setDialogPivotKey] = useState<string | null>(null);

    const openDialogForPivot = (pivotKey: string) => {
        setDialogPivotKey(pivotKey);
        setSelectedPivotKey(pivotKey);
        setIsDialogOpen(true);
    };

    const closeDialog = () => {
        setIsDialogOpen(false);
        setDialogPivotKey(null);
    };

    // Force remount of charts inside dialog when it opens (fixes google-charts render issues)
    const [dialogChartKey, setDialogChartKey] = useState<number>(0);
    React.useEffect(() => {
        if (isDialogOpen) {
            // bump key so charts remount when dialog becomes visible
            setDialogChartKey(Date.now());
        }
    }, [isDialogOpen, dialogPivotKey]);

    const resetChartsForSelection = React.useCallback(() => {
        setVisibleCharts({
            velocity: true,
            defect: true,
            rating: true,
            week: true,
            risk: true,
            critical: true,
            findings: true,
            contributors: true,
            criticalBreakdown: true,
            aging: true,
        });
        setActiveDetails(null);
    }, []);

    React.useEffect(() => {
        resetChartsForSelection();
    }, [resetChartsForSelection]);

    React.useEffect(() => {
        resetChartsForSelection();
        setIsDialogOpen(false);
        setDialogPivotKey(null);
    }, [resetChartsForSelection, selectedProjectType]);

    // convert google-chart data array -> DetailsList columns/items
    const convertDataToDetails = (data: any[][]) => {
        if (!data || !Array.isArray(data) || data.length < 1) return { columns: [], items: [] };
        const header = data[0];
        const columns: IColumn[] = header.map((h: any, i: number) => ({
            key: String(i),
            name: String(h),
            fieldName: String(i),
            minWidth: 80,
        }));
        const items = data.slice(1).map((row: any[], idx: number) => {
            const obj: any = { key: idx };
            row.forEach((cell, ci) => (obj[String(ci)] = cell));
            return obj;
        });
        return { columns, items };
    };

    const handleChartSelect = (chartKey: string, data: any[][], title: string) => {
        if (chartKey === 'Risk Summary') {
            window.open('https://ytpl.sharepoint.com/sites/qualityuat/SMARTIQPlus/SitePages/RCA-And-Raid-Logs.aspx', 'self');
        }
        if (chartKey === 'findings') {
            window.open('https://ytpl.sharepoint.com/sites/qualityuat/SMARTIQPlus/SitePages/Audit-Test.aspx', 'self');
        }
        else {
            const { columns, items } = convertDataToDetails(data);
            setVisibleCharts(prev => ({ ...prev, [chartKey]: false }));
            setActiveDetails({ key: chartKey, title, columns, items });
        }

    };

    const handleBackToChart = (chartKey: string) => {
        setVisibleCharts(prev => ({ ...prev, [chartKey]: true }));
        setActiveDetails(null);
    };

    // Wrapper: renders chart or details table based on visibleCharts/activeDetails
    const ClickableChart = (props: {
        chartKey: string;
        title: string;
        data: any[][];
        chartType: any;
        options?: any;
        width?: string;
        height?: string;
    }) => {
        const { chartKey, title, data, chartType, options, width = '100%', height = '240px' } = props;
        const visible = visibleCharts[chartKey] ?? true;
        if (!visible && activeDetails?.key === chartKey) {
            return (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <Text variant="large">{title} — Data</Text>
                        <DefaultButton text="Back to Chart" onClick={() => handleBackToChart(chartKey)} />
                    </div>
                    <DetailsList items={activeDetails.items} columns={activeDetails.columns} selectionMode={SelectionMode.none} checkboxVisibility={CheckboxVisibility.hidden} />
                </div>
            );
        }


        const safeData = Array.isArray(data) && data.length > 0 ? data : [['Category', 'Value'], ['No Data', 0]];
        return (
            <Chart
                chartType={chartType}
                width={width}
                height={height}
                data={safeData}
                options={options}
                chartEvents={[
                    {
                        eventName: 'select',
                        callback: () => {
                            handleChartSelect(chartKey, safeData, title);
                            return true;
                        },
                    },
                ]}
            />
        );
    };

    React.useEffect(() => {
        if (context && selectedProjectType) {
            loadMetricsData().catch(() => {
                // setMetricsData([]);
            });
        }
    }, [context, selectedProjectType]); // reload metrics when project type changes
    React.useEffect(() => {
        if (context) {
            loadWorkLogManagementData().catch(() => {
                setWorkLogData([]);
            });
        }
        // console.log('Selected Project Type changed:', selectedProjectType);
    }, [context, selectedProjectType]);
    React.useEffect(() => {
        if (context) {
            loadManagementEffortLogData().catch(() => {
                setManagementEffortLogData([]);
            });
        }
        // console.log('Selected Project Type changed:', selectedProjectType);
    }, [context, selectedProjectType]);
    React.useEffect(() => {
        if (context) {
            loadManagementTaskLogData().catch(() => {
                setManagementTaskLogData([]);
            });
        }
        // console.log('Selected Project Type changed:', selectedProjectType);
    }, [context, selectedProjectType]);
    React.useEffect(() => {
        if (context) {
            loadTaskManagementData().catch(() => {
                setTaskManagementData([]);
            });
        }
        // console.log('Selected Project Type changed:', selectedProjectType);
    }, [context, selectedProjectType]);

    React.useEffect(() => {
        if (context) {
            loadRAIDLogData().catch(() => { });
        }
        // console.log('Selected Project Type changed:', selectedProjectType);
    }, [context, selectedProjectType]);

    // Immediately clear dependent datasets when project type changes so async loaders
    // don't mix previous-selection data with new-selection data during fetch.
    React.useEffect(() => {
        setMetricsData([]);
        setWorkLogData([]);
        setTaskManagementData([]);
        setManagementEffortLogData([]);
        setManagementTaskLogData([]);
        setCsiData([]);
        setCSITrendData([]);
        setRAIDdata([]);
        setDefectsData([]);
        setCodeReviewDefectsData([]);
        setAMSEffortLogData([]);
        setPPOApproversData([]);
        setFilteredWorkLogData([]);
        setCodeReviewDefectDensityChartData([]);
        setPostDeliveryDefectDensityChartData([]);
        setTestingPhaseDefectDensityChartData([]);
        setscheduledVatiationChartData([]);
        setefforVatiationChartData([]);
        setOverallProductivityChartData([]);
        setResourceUtilizationTrendChartData([]);
        setCostOfQualityTrendChartData([]);
    }, [selectedProjectType]);

    const loadRAIDLogData = async () => {
        if (!context) return;
        const projectTypeAtCall = selectedProjectType;
        try {
            const genericServiceInstance: IGenericService = new GenericService(undefined, context);
            genericServiceInstance.init(undefined, context);
            const RAIDRepo: IRAIDLogRepository = new RAIDLogRepository(genericServiceInstance);
            RAIDRepo.setService(genericServiceInstance);


            // adjust call signature as needed by your repository — passing project type too
            const results = await getRiskValues(false, context);
            if (latestProjectTypeRef.current !== projectTypeAtCall) {
                console.warn('loadRAIDLogData: selection changed during fetch, discarding stale RAID data');
                return;
            }
            setRAIDdata(results || []);
            console.log('RAID data loaded', RAIDdata);
        }
        catch (err) {
            console.error('Failed to load RAID data', err);
            if (latestProjectTypeRef.current === projectTypeAtCall) {
                setRAIDdata([]);
            }
        }
    }

    React.useEffect(() => {
        if (RAIDdata.length > 0) {
            RiskChartDatafun();
        }
        // console.log('Selected Project Type changed:', selectedProjectType);
    }, [RAIDdata]);
    const RiskChartDatafun = () => {
        if (RAIDdata.length > 0) {
            //filter risk where RE >=80
            const REGEQ80 = RAIDdata.filter(i => i?.RiskExposure >= 80 && i?.SelectType == 'Risk').length
            //filter risk where RE >=60 and RE < 80
            const REGEQ60 = RAIDdata.filter(i => i?.RiskExposure >= 60 && i?.RiskExposure < 80 && i?.SelectType == 'Risk').length
            //filter risk where RE >= 0 and RE < 60
            const REGEQ0 = RAIDdata.filter(i => i?.RiskExposure >= 0 && i?.RiskExposure < 60 && i?.SelectType == 'Risk').length
            const riskChartData = [['Risk Exposure', 'Count', { role: 'annotation', type: 'string' }],
            ['RE >=80', REGEQ80, REGEQ80],
            ['RE >=60 and RE < 80', REGEQ60, REGEQ60],
            ['RE >= 0 and RE < 60', REGEQ0, REGEQ0]
            ];
            setRiskChartData(riskChartData)
        }

    }


    React.useEffect(() => {
        if (context && MetricsData.length > 0) {
            loadCSIData().catch(() => { });
        }
        // console.log('Selected Project Type changed:', selectedProjectType);
    }, [context, MetricsData]);
    // NEW: load CSI data helper — accepts optional start/end Date
    const loadCSIData = async (start?: Date | undefined, end?: Date | undefined) => {
        if (!context) return;
        const projectTypeAtCall = selectedProjectType;
        try {
            const genericServiceInstance: IGenericService = new GenericService(undefined, context);
            genericServiceInstance.init(undefined, context);
            const CSIRepo: ICSIRepository = new CSIRepository(genericServiceInstance);
            CSIRepo.setService(genericServiceInstance);


            // adjust call signature as needed by your repository — passing project type too
            const results = await getCSIValues(false, context);
            if (latestProjectTypeRef.current !== projectTypeAtCall) {
                console.warn('loadCSIData: selection changed during fetch, discarding stale CSI data');
                return;
            }
            setCsiData(results || []);
            console.log('CSI data loaded', csiData);
        } catch (err) {
            console.error('Failed to load CSI data', err);
            if (latestProjectTypeRef.current === projectTypeAtCall) {
                setCsiData([]);
            }
        }
    };
    const loadMetricsData = async () => {
        if (!context) return;
        const projectTypeAtCall = selectedProjectType;
        try {
            const genericServiceInstance: IGenericService = new GenericService(undefined, context);
            genericServiceInstance.init(undefined, context);
            const MetricsMeasurementRepo: IProjectMetricsRepository = new MetricsRepository(genericServiceInstance);
            MetricsMeasurementRepo.setService(genericServiceInstance);

            console.log('loadMetricsData: fetching metrics for projectType=', selectedProjectType);
            let MetricValues: any[] = await getMetricsFromProjectMetrics(false, context, '', selectedProjectType);

            // if (!Array.isArray(MetricValues) || MetricValues.length === 0) {
            //     console.warn('loadMetricsData: no metrics returned for projectType, retrying without projectType');
            //     MetricValues = await getMetricsFromProjectMetrics(false, context, '', undefined as any);
            // }

            const mapped = (MetricValues || []).map(m => ({
                id: (m.Metrics || '').toString(),
                title: (m.Metrics || '').toString(),
                USL: m.USL,
                LSL: m.LSL,
                goal: m.Goal,
                pivotKey: (m.Metrics || '').toString()
            }));
            const firstIndexOfKey = (arr: { id?: any; title?: any }[], key: string): number => {
                for (let i = 0; i < arr.length; i++) {
                    if (String(arr[i].id || arr[i].title || '').trim() === key) return i;
                }
                return -1;
            };
            const unique = mapped.filter((item, index, array) => {
                const key = String(item.id || item.title || '').trim();
                return key !== '' && firstIndexOfKey(array, key) === index;
            });
            if (latestProjectTypeRef.current !== projectTypeAtCall) {
                console.warn('loadMetricsData: selection changed during fetch, discarding stale metrics');
                return;
            }
            setMetricsData(unique);
            console.log('loadMetricsData: metrics loaded count=', unique.length);
        } catch (err) {
            console.error('loadMetricsData failed', err);
            if (latestProjectTypeRef.current === projectTypeAtCall) {
                setMetricsData([]);
            }
        }
    };
    const loadWorkLogManagementData = async () => {
        if (!context) return;
        const projectTypeAtCall = selectedProjectType;
        const genericServiceInstance: IGenericService = new GenericService(undefined, context);
        genericServiceInstance.init(undefined, context);
        const WorklogManagmentRepo: IWorklogManagmentRepository = new WorklogManagmentRepository(genericServiceInstance);
        WorklogManagmentRepo.setService(genericServiceInstance);

        let WorkLogValues = await getWorkLogManagementValues(false, context, selectedProjectType);
        const mapped = WorkLogValues.map(m => ({
            Title: m.Title,
            ReqTitle: m.ReqTitle,
            ProjectType: m.ProjectType,
            WorkItemNo: m.WorkItemNo,
            Simple: m.Simple,
            Medium: m.Medium,
            Complex: m.Complex,
            VeryComplex: m.VeryComplex,
            ComplexityPoints: m.ComplexityPoints,
            AppAndEnvAdjustmentFactor: m.AppAndEnvAdjustmentFactor,
            SkillAdjustmentFactor: m.SkillAdjustmentFactor,
            ReusabilityOfDesignAndCode: m.ReusabilityOfDesignAndCode,
            ExtentOfAutomation: m.ExtentOfAutomation,
            AdjustedEffort: m.AdjustedEffort,
            BaseEffort: m.BaseEffort,
            CalculatedPlannedEffort: m.CalculatedPlannedEffort,
            ActualPlannedEffort: m.ActualPlannedEffort,
            PlannedStartDate: m.PlannedStartDate,
            PlannedEndDate: m.PlannedEndDate,
            Status: m.Status,
            Remarks: m.Remarks,
            AdjustedComplexityPoint: m.AdjustedComplexityPoint,

        }));
        if (latestProjectTypeRef.current !== projectTypeAtCall) {
            console.warn('loadWorkLogManagementData: selection changed during fetch, discarding stale worklog data');
            return;
        }
        setWorkLogData(mapped);
        console.log('WorkLogData:', WorkLogData);



    };
    const loadManagementEffortLogData = async () => {
        // Implementation for loading Management Effort Log Data
        if (!context) return;
        const projectTypeAtCall = selectedProjectType;
        const genericServiceInstance: IGenericService = new GenericService(undefined, context);
        genericServiceInstance.init(undefined, context);
        const WorklogManagmentRepo: IManagemnetEffortLogRepository = new ManagemnetEffortLogRepository(genericServiceInstance);
        WorklogManagmentRepo.setService(genericServiceInstance);

        let WorkLogValues = await getManagementEffortLogValues(false, context, selectedProjectType);
        const mapped = WorkLogValues.map(m => ({
            Title: m.Title,
            TaskDescription: m?.TaskDescription,
            ActualStartDate: m?.ActualStartDate,
            ActualEndDate: m?.ActualEndDate,
            ActualEffortHrs: m?.ActualEffortHrs,
            UpdatedBy: m?.UpdatedBy,
            Remarks: m?.Remarks,
            ManagementTaskActivity: m?.ManagementTaskActivity,
            Completion: m?.Completion,

        }));
        if (latestProjectTypeRef.current !== projectTypeAtCall) {
            console.warn('loadManagementEffortLogData: selection changed during fetch, discarding stale management effort data');
            return;
        }
        setManagementEffortLogData(mapped);
        console.log('ManagementEffortLogData:', ManagementEffortLogData);
    };
    const loadManagementTaskLogData = async () => {
        // Implementation for loading Management Task Log Data
        if (!context) return;
        const projectTypeAtCall = selectedProjectType;
        const genericServiceInstance: IGenericService = new GenericService(undefined, context);
        genericServiceInstance.init(undefined, context);
        const WorklogManagmentRepo: ITaskManagementLogRepository = new TaskManagementLogRepository(genericServiceInstance);
        WorklogManagmentRepo.setService(genericServiceInstance);

        let WorkLogValues = await getTaskManagementLogValues(false, context, selectedProjectType);
        const mapped = WorkLogValues.map(m => ({
            Title: m.Title,
            Activity: m.Activity,
            Responsibility: m.Responsibility,
            TaskDescription: m.TaskDescription,
            PlannedStartDate: m.PlannedStartDate,
            PlannedEndDate: m.PlannedEndDate,
            PlannedEffortHrs: m.PlannedEffortHrs,
            TaskStatus: m.TaskStatus,
            Phase: m.Phase,
            Remarks: m.Remarks,
            Completion: m.Completion,
            CompletionCount: m.CompletionCount,


        }));
        if (latestProjectTypeRef.current !== projectTypeAtCall) {
            console.warn('loadManagementTaskLogData: selection changed during fetch, discarding stale management task data');
            return;
        }
        setManagementTaskLogData(mapped);
        console.log('ManagementTaskLogData:', ManagementTaskLogData);
    };
    const loadTaskManagementData = async () => {
        // Implementation for loading Management Task Log Data
        if (!context) return;
        const projectTypeAtCall = selectedProjectType;
        const genericServiceInstance: IGenericService = new GenericService(undefined, context);
        genericServiceInstance.init(undefined, context);
        const WorklogManagmentRepo: ITaskManagementRepository = new TaskManagementRepository(genericServiceInstance);
        WorklogManagmentRepo.setService(genericServiceInstance);

        let WorkLogValues = await getTaskManagementValues(false, context, selectedProjectType);
        const mapped = WorkLogValues.map(m => ({
            WorkItemNo: m.WorkItemNo,
            ReqTitle: m.ReqTitle,
            ActualEffort: m.ActualEffort,
            ActualEndDate: m.ActualEndDate,
            ActualStartDate: m.ActualStartDate,
            AssignedTo: m.AssignedTo,
            Remarks: m.Remarks,
            TaskStatus: m.TaskStatus,
            TaskType: m.TaskType,


        }));
        if (latestProjectTypeRef.current !== projectTypeAtCall) {
            console.warn('loadTaskManagementData: selection changed during fetch, discarding stale task data');
            return;
        }
        setTaskManagementData(mapped);
        console.log('TaskManagementData:', TaskManagementData);
    };
    React.useEffect(() => {
        PPOApprovers().catch(() => {
            //setPPOApproversData([]);
        });
    }, [context]);
    const PPOApprovers = async () => {
        if (!context) return;
        const projectTypeAtCall = selectedProjectType;
        try {
            const genericServiceInstance: IGenericService = new GenericService(undefined, context);
            genericServiceInstance.init(undefined, context);
            const PPOApproversRepo: IPPOApproversRepository = new PPOApproversRepository(genericServiceInstance);
            PPOApproversRepo.setService(genericServiceInstance);
            const results = await getPPOApproversValues(false, context);
            // optional: dedupe/sort depending on repo shape — store as-is
            if (results.length > 0) {
                console.log('PPOApprovers Results', results);
                if (latestProjectTypeRef.current !== projectTypeAtCall) {
                    console.warn('PPOApprovers: selection changed during fetch, discarding stale PPO approver data');
                    return;
                }
                setPPOApproversData(results);
            }

            console.log('PPOApprovers data loaded', PPOApproversData);
        } catch (err) {
            console.error('Failed to load PPOApprovers data', err);

        }
    }






    function normalizeToLocalDateOnly(input: any): Date | null {
        const d = input instanceof Date ? input : new Date(input);
        if (isNaN(d.getTime())) return null;
        // Strip time to local midnight
        return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }

    {/* -------------------------------- DEVM Dashboard Data calculations started ---------------------- */ }
    React.useEffect(() => {
        // Run calculations on any data change; calculations handle empty arrays

        console.log('metricsData:', MetricsData);
        if (WorkLogData.length > 0 && TaskManagementData.length > 0 && selectedProjectType && MetricsData.length > 0) {


            // Create a quick lookup: closed tasks by WorkItemNo
            //let closedTaskByWorkItem: any[] = [];
            const closedWorkLogByWorkItem = WorkLogData.filter(wl => {
                const status = String(wl?.Status ?? '').toLowerCase();
                const projectType = String(wl?.ProjectType ?? '').toLowerCase();
                const matchesStatus = status === 'completed' || status === 'closed';
                const matchesProjectType = projectType === String(selectedProjectType ?? '').toLowerCase();
                return matchesStatus && matchesProjectType && wl?.WorkItemNo != null;
                //&& closedTaskByWorkItem.some(t => t.WorkItemNo === wl.WorkItemNo)
            });
            const closedTaskByWorkItem = TaskManagementData.filter(t => ((t?.TaskStatus || '').toLowerCase() === 'completed' || (t?.TaskStatus || '').toLowerCase() === 'closed')
                && t?.WorkItemNo != null && // && new Date(t.ActualEndDate).getMonth().toString() === selectedMonth 
                closedWorkLogByWorkItem.some(wl => wl.WorkItemNo === t.WorkItemNo));
            console.log('Filtered CompletedTasks result:', closedTaskByWorkItem);
            //calcuating Planned Duration



            // Precompute totals per WorkItemNo (in-scope, no helper functions)
            const totalsByWorkItem = new Map<string, number>();
            const totalEffortForRequirementsAnalysisByWorkItem = new Map<string, number>();
            const totalEffortForCodingByWorkItem = new Map<string, number>();
            const totalEffortForCodeReviewByWorkItem = new Map<string, number>();
            const totalEffortForCodeRework = new Map<string, number>();
            const totalEffortForUnitTesting = new Map<string, number>();
            const totalEffortForTestingExecution = new Map<string, number>();
            const latestEndDateByWorkItem = new Map<string, Date>();


            // track latest end date
            for (const t of closedTaskByWorkItem) {
                const key = String(t?.WorkItemNo ?? '').trim();
                if (key !== '') {
                    const eff = Number(t?.ActualEffort) || 0;
                    totalsByWorkItem.set(key, (totalsByWorkItem.get(key) ?? 0) + eff);
                    const tt = (t?.TaskType || '').toLowerCase();
                    totalEffortForRequirementsAnalysisByWorkItem.set(key, (totalEffortForRequirementsAnalysisByWorkItem.get(key) ?? 0) + (tt === 'requirements analysis' ? eff : 0));
                    totalEffortForCodingByWorkItem.set(key, (totalEffortForCodingByWorkItem.get(key) ?? 0) + (tt === 'coding' ? eff : 0));
                    totalEffortForCodeReviewByWorkItem.set(key, (totalEffortForCodeReviewByWorkItem.get(key) ?? 0) + (tt === 'code review' ? eff : 0));
                    totalEffortForCodeRework.set(key, (totalEffortForCodeRework.get(key) ?? 0) + (tt === 'rework after code review' ? eff : 0));
                    totalEffortForUnitTesting.set(key, (totalEffortForUnitTesting.get(key) ?? 0) + (tt === 'unit testing' ? eff : 0));
                    totalEffortForTestingExecution.set(key, (totalEffortForTestingExecution.get(key) ?? 0) + (tt === 'system testing' || tt === 'integration testing' ? eff : 0));

                    const rawEnd = t?.ActualEndDate; // could be string or Date

                    if (rawEnd != null && rawEnd !== '') {
                        const end = rawEnd instanceof Date ? rawEnd : new Date(rawEnd);
                        if (!isNaN(end.getTime())) {
                            const prev = latestEndDateByWorkItem.get(key);
                            if (!prev || end > prev) {
                                latestEndDateByWorkItem.set(key, end);
                            }
                        }
                    }

                }
            }
            console.log('Totals by WorkItemNo:', totalsByWorkItem);

            //Calcuting the Actual Efforts of the tasks based on the work items
            const WorkItemActualEffort: any[] = [];
            const seen = new Set<string>();
            closedTaskByWorkItem.forEach(task => {
                const hasWorkItemNo =
                    task?.WorkItemNo != null &&
                    task?.WorkItemNo !== undefined &&
                    String(task.WorkItemNo).trim() !== '';

                if (hasWorkItemNo) {
                    const key = String(task.WorkItemNo).trim();
                    const totalActualEffort = totalsByWorkItem.get(key) ?? 0;
                    const ActualEndDate = latestEndDateByWorkItem.get(key) ?? '';
                    // skip if the work item is seen
                    if (seen.has(key)) return;


                    WorkItemActualEffort.push({
                        WorkItemNo: task.WorkItemNo,
                        ActualEffort: totalActualEffort,
                        ActualEndDate: ActualEndDate,
                        reqAnalysisEffort: totalEffortForRequirementsAnalysisByWorkItem.get(key) ?? 0,
                        codingEffort: totalEffortForCodingByWorkItem.get(key) ?? 0,
                        codeReviewEffort: totalEffortForCodeReviewByWorkItem.get(key) ?? 0,
                        codeReworkEffort: totalEffortForCodeRework.get(key) ?? 0,
                        unitTestingEffort: totalEffortForUnitTesting.get(key) ?? 0,
                        testingExecutionEffort: totalEffortForTestingExecution.get(key) ?? 0,

                    });
                    // task.ActualEffort = totalActualEffort;

                    seen.add(key);
                    console.log(`Updated ActualEffort for WorkItemNo ${task.WorkItemNo}:`, totalActualEffort);
                } else {
                    task.ActualEffort = task.ActualEffort; // keep original
                }

            });

            // get the planned effort from workLog Data
            const WorkLogItemWithPlannedandActualEfforts: any[] = []
            let EffortVariation = 0.0;
            let meanEffortVariation = 0.0;
            let meanSV = 0.0;
            let diffDaysSV = 0;
            let PlannedDuration = 0;
            let overAllProductivity = 0.0;

            closedWorkLogByWorkItem.forEach(worklog => {
                if (!worklog?.WorkItemNo) {
                    return;
                }
                WorkItemActualEffort.filter(item => item.WorkItemNo === worklog.WorkItemNo).forEach(
                    i => {
                        EffortVariation = Math.round((i.ActualEffort - worklog.ActualPlannedEffort) * 100 / worklog.ActualPlannedEffort);

                        const actualEndDate = normalizeToLocalDateOnly(i.ActualEndDate);
                        const plannedEndDate = normalizeToLocalDateOnly(worklog.PlannedEndDate);
                        const PlannedStartDate = normalizeToLocalDateOnly(worklog.PlannedStartDate);


                        if (actualEndDate && plannedEndDate) {
                            const diffMs = actualEndDate.getTime() - plannedEndDate.getTime();
                            diffDaysSV = Math.round((diffMs / (1000 * 60 * 60 * 24)) * 100) / 100; // 2 decimals
                            console.log('Days difference:', diffDaysSV);
                        }

                        if (plannedEndDate && PlannedStartDate) {
                            const diffMs = plannedEndDate.getTime() - PlannedStartDate.getTime();
                            PlannedDuration = (Math.round((diffMs / (1000 * 60 * 60 * 24)) * 100) / 100) + 1; // 2 decimals
                            console.log('Days difference:', diffDaysSV);
                        }



                        //const PlannedDuration = (new Date(worklog.PlannedEndDate).getTime() - new Date(worklog.PlannedStartDate).getTime()) + 1
                        const ScheduledVariation = diffDaysSV * 100 / PlannedDuration
                        overAllProductivity = i.ActualEffort / worklog.AdjustedComplexityPoint

                        let ReqAnalyisEffortDensity = i.reqAnalysisEffort / worklog.AdjustedComplexityPoint;
                        let codingProductivity = i.codingEffort / worklog.AdjustedComplexityPoint;
                        let CodeReviewEffortDensity = i.codeReviewEffort / worklog.AdjustedComplexityPoint;
                        let UnitTestingEffortDensity = i.unitTestingEffort / worklog.AdjustedComplexityPoint;
                        let TestingExecutionEffortDensity = i.testingExecutionEffort / worklog.AdjustedComplexityPoint;
                        let CodeReworkEffortDensity = i.codeReworkEffort / worklog.AdjustedComplexityPoint;

                        // closedTaskByWorkItem.forEach(task => {
                        // if (i.TaskType.toLowerCase() === 'requirements analysis' && worklog.AdjustedComplexityPoint > 0) {
                        //     ReqAnalyisEffortDensity += (i.ActualEffort / worklog.AdjustedComplexityPoint);
                        // }
                        // if ( i.TaskType.toLowerCase() === 'coding' && worklog.AdjustedComplexityPoint > 0) {
                        //     codingProductivity += (i.ActualEffort / worklog.AdjustedComplexityPoint);
                        // }
                        // if(i.TaskType.toLowerCase()==='code review' && worklog.AdjustedComplexityPoint > 0){
                        //     CodeReviewEffortDensity += (i.ActualEffort / worklog.AdjustedComplexityPoint);
                        // }
                        // });

                        WorkLogItemWithPlannedandActualEfforts.push({
                            ...i, PlannedEffort: worklog.ActualPlannedEffort,
                            EffortVariation: EffortVariation,
                            PlannedEndDate: worklog.PlannedEndDate.toLocaleString(),
                            PlannedStartDate: worklog.PlannedStartDate.toLocaleString(),
                            PlannedDuration: PlannedDuration,
                            ScheduledVariation: ScheduledVariation,
                            size: worklog.AdjustedComplexityPoint,
                            OverAllProductivity: overAllProductivity,
                            ReqAnalyisEffortDensity: ReqAnalyisEffortDensity,
                            codingProductivity: codingProductivity,
                            CodeReviewEffortDensity: CodeReviewEffortDensity,
                            CodeReworkEffortDensity: CodeReworkEffortDensity,
                            UnitTestingEffortDensity: UnitTestingEffortDensity,
                            TestingExecutionEffortDensity: TestingExecutionEffortDensity,
                        });
                        console.log(`Updated PlannedEffort for WorkItemNo `, WorkLogItemWithPlannedandActualEfforts);
                    }
                );

            });

            //DevM Dashboard Monthly Data calculations
            const groupedByMonth: { [key: string]: MonthlyData } = {};
            WorkLogItemWithPlannedandActualEfforts.forEach(item => {
                const actualEndDate = normalizeToLocalDateOnly(item.ActualEndDate);
                if (actualEndDate) {
                    const monthKey = `${actualEndDate.getFullYear()}-${actualEndDate.getMonth() + 1}`;
                    //Adding the Actual Effort for the work items whose Actual End Date is in the same month from WorkLogItemWithPlannedandActualEfforts;
                    if (!groupedByMonth[monthKey]) {
                        groupedByMonth[monthKey] = {
                            ActualEffort: 0,
                            PlannedEffort: 0,
                            size: 0,
                            plannedDuration: 0,
                            codeReviewDefects: 0,

                        };
                    }

                    groupedByMonth[monthKey].ActualEffort += Number(item.ActualEffort) || 0;
                    groupedByMonth[monthKey].PlannedEffort += Number(item.PlannedEffort) || 0;
                    groupedByMonth[monthKey].size += Number(item.size) || 0;
                    groupedByMonth[monthKey].plannedDuration += Number(item.PlannedDuration) || 0;
                    // groupedByMonth[monthKey].codeReviewDefects += CodeReviewDefectsData.filter(crD => {
                    //     const defectDate = normalizeToLocalDateOnly(crD.IdentifiedDate);
                    //     return defectDate && defectDate.getFullYear() === actualEndDate.getFullYear() && defectDate.getMonth() === actualEndDate.getMonth();
                    // }).length;
                }
            });

            const workItemsForDefectDensity = new Set<string>();
            const workItemLookup = new Map<string, string>();
            const complexityByWorkItem = new Map<string, number>();
            WorkLogItemWithPlannedandActualEfforts.forEach(item => {
                const key = String(item?.WorkItemNo ?? '').trim();
                if (!key) {
                    return;
                }
                workItemsForDefectDensity.add(key);
                workItemLookup.set(key.toLowerCase(), key);
                const rawSize = item?.size ?? item?.AdjustedComplexityPoint ?? 0;
                const sizeValue = Number(rawSize);
                if (!isNaN(sizeValue)) {
                    complexityByWorkItem.set(key, sizeValue);
                }
            });
            const codeReviewDefectsByMonth = new Map<string, number>();
            (CodeReviewDefectsData || []).forEach(defect => {
                const requirementIdRaw = String(defect?.RequirementID ?? '').trim();
                if (!requirementIdRaw) {
                    return;
                }
                const resolvedRequirementKey = workItemsForDefectDensity.has(requirementIdRaw)
                    ? requirementIdRaw
                    : workItemLookup.get(requirementIdRaw.toLowerCase());
                if (!resolvedRequirementKey) {
                    return;
                }
                const defectDate = normalizeToLocalDateOnly(defect?.IdentifiedDate || defect?.ReviewCompletionDate);
                if (!defectDate) {
                    return;
                }
                const key = `${defectDate.getFullYear()}-${defectDate.getMonth() + 1}`;
                codeReviewDefectsByMonth.set(key, (codeReviewDefectsByMonth.get(key) ?? 0) + 1);
            });
            Object.keys(groupedByMonth).forEach(monthKey => {
                groupedByMonth[monthKey].codeReviewDefects = codeReviewDefectsByMonth.get(monthKey) ?? 0;
            });


            // let sumOfEV = 0.0
            let standardDeviationEV = 0.0;
            let standardDeviationSV = 0.0;
            let meanOverallProductivity = 0.0;


            // Calculating Mean and Standard Deviation for Effort Variation, Scheduled Variation and Overall Productivity
            const filteredForMonth = WorkLogItemWithPlannedandActualEfforts.filter(i => {
                try {
                    const d = i && i.ActualEndDate ? new Date(i.ActualEndDate) : null;
                    return d && d.getMonth().toString() === selectedMonth;
                } catch (e) {
                    return false;
                }
            });
            const n = filteredForMonth.length;
            if (n > 0) {
                meanEffortVariation = filteredForMonth.reduce((sum, x) => sum + (Number(x.EffortVariation) || 0), 0) / n;
                meanSV = filteredForMonth.reduce((sum, x) => sum + (Number(x.ScheduledVariation) || 0), 0) / n;
                meanOverallProductivity = filteredForMonth.reduce((sum, x) => sum + (Number(x.OverAllProductivity) || 0), 0) / n;
                const varianceEV = filteredForMonth.reduce((acc, x) => acc + ((Number(x.EffortVariation) || 0) - meanEffortVariation) ** 2, 0) / n;
                const varianceSV = filteredForMonth.reduce((acc, x) => acc + ((Number(x.ScheduledVariation) || 0) - meanSV) ** 2, 0) / n;
                standardDeviationEV = parseFloat(Math.sqrt(varianceEV).toFixed(2));
                standardDeviationSV = parseFloat(Math.sqrt(varianceSV).toFixed(2));
                setMeanEffortVariation(parseFloat(meanEffortVariation.toFixed(2)));
                setMeanSV(parseFloat(meanSV.toFixed(2)));
                setMeanOverallProductivity(parseFloat(meanOverallProductivity.toFixed(2)));
                setStandardDeviationOfEV(standardDeviationEV);
                setStandardDeviationOfSV(standardDeviationSV);
            } else {
                setMeanEffortVariation(0);
                setMeanSV(0);
                setMeanOverallProductivity(0);
                setStandardDeviationOfEV(0);
                setStandardDeviationOfSV(0);
            }
            const OAPSize = WorkLogItemWithPlannedandActualEfforts.filter(i => new Date(i.ActualEndDate).getMonth().toString() == selectedMonth).reduce((sum, x) => sum + x.size, 0);
            const OAPEffort = WorkLogItemWithPlannedandActualEfforts.filter(i => new Date(i.ActualEndDate).getMonth().toString() == selectedMonth).reduce((sum, x) => sum + x.ActualEffort, 0);
            setOverAllProductivitySize(OAPSize);
            setActualEffortOverAllProductivity(OAPEffort);
            const numofdefects = defectsData.filter(dD => dD.DefectDetectedPhase == 'Post Production' && new Date(dD.DefectDetectedOn).getMonth().toString() == selectedMonth).length
            setDefectCount(numofdefects);
            //create a chart for Defect Density


            //Calculating the number of Code Review Defects for the selected month and for completed task in worklog managemnet;
            let totalCodeReviewDefectsForSelectedMonth = 0;
            WorkLogItemWithPlannedandActualEfforts.filter(wl => new Date(wl.ActualEndDate).getMonth().toString() == selectedMonth).forEach(wlItem => {
                const codeReviewDefectsForWorkItem = CodeReviewDefectsData.filter(crD => crD.RequirementID == wlItem.WorkItemNo && new Date(crD.IdentifiedDate).getMonth().toString() == selectedMonth).length
                totalCodeReviewDefectsForSelectedMonth += codeReviewDefectsForWorkItem;
            });
            setCodeReviewDefectCount(totalCodeReviewDefectsForSelectedMonth);


            WorkLogItemWithPlannedandActualEfforts.filter(wl => new Date(wl.ActualEndDate).getMonth().toString() == selectedMonth).forEach(wlItem => {
                const codeReviewDefectsForWorkItem = CodeReviewDefectsData.filter(crD => crD.RequirementID == wlItem.WorkItemNo && new Date(crD.IdentifiedDate).getMonth().toString() == selectedMonth).length
                totalCodeReviewDefectsForSelectedMonth += codeReviewDefectsForWorkItem;
            });
            setCodeReviewDefectCount(totalCodeReviewDefectsForSelectedMonth);


            // const variance = nums.reduce((acc, x) => acc + (x - mean) ** 2, 0) / n;
            //return Math.sqrt(variance);

            //Create a chart for Effort Distribution
            let ManagementEffortLogDataFiltered = ManagementEffortLogData.filter(m => m?.ActualEndDate != null && new Date(m.ActualEndDate).getMonth().toString() === selectedMonth && m.Completion).reduce((sum, x) => sum + x.ActualEffortHrs, 0);
            console.log('ManagementEffortLogDataFiltered:', ManagementEffortLogDataFiltered);
            let effortDistributionData: any[] = [];
            effortDistributionData.push(['Task Type', 'Effort Hrs']);
            effortDistributionData.push(['Management Efforts', ManagementEffortLogDataFiltered]);
            // Aggregate Actual Efforts from closedTaskByWorkItem by TaskType
            const effortByType = new Map<string, number>();
            if (Array.isArray(closedTaskByWorkItem)) {
                closedTaskByWorkItem.forEach(item => {
                    const type = (item?.TaskType ?? 'Unspecified').toString();
                    const val = Number(item?.ActualEffort) || 0;
                    effortByType.set(type, (effortByType.get(type) || 0) + val);
                });
            }
            effortByType.forEach((sum, type) => {
                effortDistributionData.push([type, sum]);
            });
            setEffortDistributionData(effortDistributionData);

            //Resource utilization for selected month
            //let resourceUtilizationData = 0;
            let TotalActualEffort = 0;
            closedTaskByWorkItem.forEach(item => {
                if (new Date(item.ActualStartDate).getMonth().toString() === selectedMonth) {
                    TotalActualEffort += item.ActualEffort
                }

            });
            ManagementEffortLogData.forEach(mE => {
                if (new Date(mE.ActualStartDate).getMonth().toString() === selectedMonth) {
                    TotalActualEffort += mE.ActualEffortHrs
                }
            });
            AMSEffortLogData.forEach(aE => {
                //console.log('aE.ActualStartDate:', aE.ActualStartDate);
                if (new Date(aE.ActualStartDate).getMonth().toString() === selectedMonth) {
                    TotalActualEffort += aE.ActualEfforts
                }
            });

            //Calculating total working days in the selected month
            let totalWorkingDaysInMonth = 19; //Assuming 8 hours per day for 24 working days

            //const totalWorkingDaysInMonth = getWorkingDaysInMonth(new Date().getFullYear(), Number(selectedMonth));
            const totalAvailableHours = totalWorkingDaysInMonth * 8;
            let teamSize = 1;
            if (PPOApproversData != undefined && PPOApproversData.length > 0) {
                PPOApproversData.forEach(pA => {
                    if (pA.SiteURL == context?.pageContext.web.absoluteUrl) {
                        teamSize = Number(pA.teamSize);
                    }
                })
            }
            let AllocatedEffort = Number(teamSize) * totalAvailableHours;
            let resourceUtilization = (TotalActualEffort / AllocatedEffort) * 100;
            setResourceUtilization(parseFloat(resourceUtilization.toFixed(2)));

            const monthlyEffort = new Map<string, number>();
            const addEffort = (raw: any, val: number) => {
                if (raw == null || val == null) return;
                const d = raw instanceof Date ? raw : new Date(raw);
                if (isNaN(d.getTime())) return;
                const key = `${d.getFullYear()}-${('0' + (d.getMonth() + 1)).slice(-2)}`;
                monthlyEffort.set(key, (monthlyEffort.get(key) || 0) + Number(val));
            };
            closedTaskByWorkItem.forEach(item => addEffort(item?.ActualStartDate, item?.ActualEffort));
            ManagementEffortLogData.forEach(m => addEffort(m?.ActualStartDate, m?.ActualEffortHrs));
            AMSEffortLogData.forEach(a => addEffort(a?.ActualStartDate, a?.ActualEfforts));
            const monthlyWorkingDays = 19;
            const monthlyAvailableHours = monthlyWorkingDays * 8;
            let monthlyTeamSize = 1;
            if (Array.isArray(PPOApproversData)) {
                PPOApproversData.forEach(p => {
                    if (p?.SiteURL === context?.pageContext.web.absoluteUrl) {
                        const size = Number(p.teamSize ?? p.TeamSize);
                        if (!isNaN(size) && size > 0) monthlyTeamSize = size;
                    }
                });
            }
            const monthlyAllocation = monthlyAvailableHours * monthlyTeamSize;
            const trendRows: any[] = [['Month', 'Resource Utilization (%)', { role: 'annotation', type: 'string' }, 'LSL', 'USL']];
            const monthlyEffortKeys: string[] = [];
            monthlyEffort.forEach((_, k) => {
                monthlyEffortKeys.push(k);
            });
            let USLLSLValuesRU = MetricsData.filter(m => m.title === 'Resource Utilization').map(i => {
                return { USL: i.USL, LSL: i.LSL, Title: i.Title }
            });
            monthlyEffortKeys.sort((a, b) => a.localeCompare(b)).forEach(key => {
                const sum = monthlyEffort.get(key) || 0;
                const value = monthlyAllocation === 0 ? 0 : (sum / monthlyAllocation) * 100;
                trendRows.push([key, parseFloat(value.toFixed(2)), parseFloat(value.toFixed(2)).toString(), USLLSLValuesRU[0]?.LSL ?? 0, USLLSLValuesRU[0]?.USL ?? 0]);
            });
            setResourceUtilizationTrendChartData(trendRows);

            //To create a cost of quality we need Prevention cost, appraisal cost,failure cost
            //TotalActualEffort
            //    const failureCost = (() => {
            //        const toNum = (v: any) => {
            //            const n = Number(v);
            //            return isNaN(n) ? 0 : n;
            //        };
            let failureCost = 0;

            // 1) closedTaskByWorkItem: TaskType starts with "Rework" OR equals "Re-testing"/"retesting"
            if (Array.isArray(closedTaskByWorkItem)) {
                closedTaskByWorkItem.forEach(item => {
                    if (new Date(item?.ActualStartDate).getMonth().toString() === selectedMonth) {
                        const tt = String(item?.TaskType ?? '').trim().toLowerCase();
                        if ((tt.toLocaleLowerCase().indexOf('rework') === 0) || tt.toLocaleLowerCase() === 're-testing' || tt.toLocaleLowerCase() === 'retesting') {
                            failureCost += Number(item?.ActualEffort);
                        }
                    }
                });
            }

            // 2) ManagementEffortLogData: ManagementTaskActivity starts with "Rework" OR equals specified implementations
            if (Array.isArray(ManagementEffortLogData)) {
                ManagementEffortLogData.forEach(m => {
                    // tolerate possible property name variations
                    if (new Date(m?.ActualStartDate).getMonth().toString() === selectedMonth) {
                        const activity = String(m?.ManagementTaskActivity ?? m?.ManagmentTaskActivity ?? '').trim().toLowerCase();
                        if (
                            (activity.toLocaleLowerCase().indexOf('rework') === 0) ||
                            activity === 'implementation of corrective action' ||
                            activity === 'implementation of preventive action'
                        ) {
                            failureCost += Number(m?.ActualEffortHrs);
                        }
                    }
                });
            }

            // 3) AMSEffortLogData: TaskType equals "solution rework effort"
            if (Array.isArray(AMSEffortLogData)) {
                AMSEffortLogData.forEach(a => {
                    if (new Date(a?.ActualStartDate).getMonth().toString() === selectedMonth) {
                        const at = String(a?.TaskType ?? '').trim().toLowerCase();
                        if (at.toLocaleLowerCase() === 'solution rework effort') {
                            failureCost += Number(a?.ActualEfforts);
                        }
                    }
                });
            }
            //Appraisal Cost
            //    const appraisalCost = (() => {
            //        const toNum = (v: any) => {
            //            const n = Number(v);
            //            return isNaN(n) ? 0 : n;
            //        };
            let appraisalCost = 0;

            // closedTaskByWorkItem: exact matches for testing/review-related types
            const appraisalTypes = new Set([
                'unit testing', 'code review', 'system testing', 'integration testing', 'acceptance testing',
                'review of test cases', 'review of design document', 'review of unit test cases', 'review of release notes',
                'review build', 'review of test plan', 'review of test environment setup', 'review delivery notes'
            ].map(s => s.toLowerCase()));

            if (Array.isArray(closedTaskByWorkItem)) {
                closedTaskByWorkItem.forEach(item => {
                    if (new Date(item?.ActualStartDate).getMonth().toString() !== selectedMonth) return;
                    const tt = String(item?.TaskType ?? '').trim().toLowerCase();
                    if (appraisalTypes.has(tt)) appraisalCost += Number(item?.ActualEffort);
                });
            }

            // ManagementEffortLogData: ManagementTaskActivity == 'Audits'
            if (Array.isArray(ManagementEffortLogData)) {
                ManagementEffortLogData.forEach(m => {
                    if (new Date(m?.ActualStartDate).getMonth().toString() !== selectedMonth) return;
                    const activity = String(m?.ManagementTaskActivity ?? m?.ManagmentTaskActivity ?? '').trim().toLowerCase();
                    if (activity === 'audits' || activity === 'audit') appraisalCost += Number(m?.ActualEffortHrs);
                });
            }

            // AMSEffortLogData: TaskType matches solution review/testing (tolerate 'soultion' spelling)

            if (Array.isArray(AMSEffortLogData)) {
                AMSEffortLogData.forEach(a => {
                    if (new Date(a?.ActualStartDate).getMonth().toString() !== selectedMonth) return;
                    const at = String(a?.TaskType ?? '').trim().toLowerCase();
                    if (at === 'solution review effort' || at === 'solution testing effort' || at === 'soultion review effort' || at === 'soultion testing effort') {
                        appraisalCost += Number(a?.ActualEfforts);
                    }
                });
            }


            //Prevention Cost
            let preventionCost = 0;
            preventionCost = ManagementEffortLogDataFiltered
            let costOfQualityTotal = (preventionCost + appraisalCost + failureCost) * 100 / TotalActualEffort;

            setcostOfQualityTotal(parseFloat(costOfQualityTotal.toFixed(2)));
            const monthKeyFrom = (raw: any) => {
                if (!raw) return null;
                const d = raw instanceof Date ? raw : new Date(raw);
                if (isNaN(d.getTime())) return null;
                return `${d.getFullYear()}-${('0' + (d.getMonth() + 1)).slice(-2)}`;
            };
            const monthlyBuckets = new Map<string, { prevention: number; appraisal: number; failure: number }>();
            const ensureBucket = (key: string | null) => {
                if (!key) return null;
                if (!monthlyBuckets.has(key)) {
                    monthlyBuckets.set(key, { prevention: 0, appraisal: 0, failure: 0 });
                }
                return monthlyBuckets.get(key)!;
            };
            closedTaskByWorkItem.forEach(item => {
                const bucket = ensureBucket(monthKeyFrom(item?.ActualStartDate));
                if (!bucket) return;
                const tt = String(item?.TaskType ?? '').trim().toLowerCase();
                if (appraisalTypes.has(tt)) bucket.appraisal += Number(item?.ActualEffort) || 0;
                if (tt.indexOf('rework') || tt === 're-testing' || tt === 'retesting') bucket.failure += Number(item?.ActualEffort) || 0;
            });
            ManagementEffortLogData.forEach(m => {
                const bucket = ensureBucket(monthKeyFrom(m?.ActualStartDate));
                if (!bucket) return;
                const activity = String(m?.ManagementTaskActivity ?? m?.ManagmentTaskActivity ?? '').trim().toLowerCase();
                const effort = Number(m?.ActualEffortHrs) || 0;
                if (activity === 'audits' || activity === 'audit') bucket.appraisal += effort;
                if (activity.indexOf('rework') || activity === 'implementation of corrective action' || activity === 'implementation of preventive action') bucket.failure += effort;
                if (m?.Completion) bucket.prevention += effort;
            });
            AMSEffortLogData.forEach(a => {
                const bucket = ensureBucket(monthKeyFrom(a?.ActualStartDate));
                if (!bucket) return;
                const at = String(a?.TaskType ?? '').trim().toLowerCase();
                const effort = Number(a?.ActualEfforts) || 0;
                if (at === 'solution review effort' || at === 'solution testing effort' || at === 'soultion review effort' || at === 'soultion testing effort') bucket.appraisal += effort;
                if (at === 'solution rework effort') bucket.failure += effort;
            });
            const sortedMonths: string[] = [];
            monthlyBuckets.forEach((_, k) => {
                sortedMonths.push(k);
            });
            sortedMonths.sort((a, b) => a.localeCompare(b));
            let USLLSLValuesCOQ = MetricsData.filter(m => m.title === 'Cost of Quality').map(i => {
                return { USL: i.USL, LSL: i.LSL, Title: i.Title }
            });
            const coqTrendRows: any[] = [['Month', 'Cost of Quality (%)', { role: 'annotation', type: 'string' }, 'LSL', 'USL']];
            sortedMonths.forEach(monthKey => {
                const bucket = monthlyBuckets.get(monthKey)!;
                const denominator = (monthlyEffort.get(monthKey) || 0) || 1;
                const totalCost = bucket.prevention + bucket.appraisal + bucket.failure;
                coqTrendRows.push([
                    monthKey,
                    parseFloat(((totalCost * 100) / denominator).toFixed(2)),
                    parseFloat(((totalCost * 100) / denominator).toFixed(2)),
                    USLLSLValuesCOQ[0]?.LSL ?? 0,
                    USLLSLValuesCOQ[0]?.USL ?? 0
                ]);
            });
            setCostOfQualityTrendChartData(coqTrendRows);
            //code review defects list unittesting from testing defects data systemtesting defects from testing defects data
            testingDefectDataForDashboard(WorkItemActualEffort);

             const codeReviewMetricBounds = MetricsData.filter(m => {
                    const title = String(m.title ?? '').trim().toLowerCase();
                    return title === 'defect density_month' || title === 'defect density' || title === 'internal defect density';
                }).map(i => ({ USL: i.USL, LSL: i.LSL }));
                const codeReviewLsl = codeReviewMetricBounds[0]?.LSL ?? 0;
                const codeReviewUsl = codeReviewMetricBounds[0]?.USL ?? 0;
                const codeReviewDefectDensityTrendData: any[] = [['Month', 'Code Review Defect Density', { role: 'annotation', type: 'string' }, 'LSL', 'USL']];
                Object.keys(groupedByMonth).sort((a, b) => a.localeCompare(b)).forEach(monthKey => {
                    const monthAggregate = groupedByMonth[monthKey];
                    const density = monthAggregate.size > 0 ? monthAggregate.codeReviewDefects / monthAggregate.size : 0;
                    const roundedDensity = parseFloat(density.toFixed(3));
                    codeReviewDefectDensityTrendData.push([
                        monthKey,
                        roundedDensity,
                        roundedDensity.toFixed(3),
                        codeReviewLsl ?? 0,
                        codeReviewUsl ?? 0
                    ]);
                });
                if (codeReviewDefectDensityTrendData.length === 1) {
                    codeReviewDefectDensityTrendData.push([
                        'No Data',
                        0,
                        '0.000',
                        codeReviewLsl ?? 0,
                        codeReviewUsl ?? 0
                    ]);
                }
                console.log('codeReviewDefectDensityTrendData:', codeReviewDefectDensityTrendData);
                setCodeReviewDefectDensityChartData(codeReviewDefectDensityTrendData);

                const deliveredDefectBounds = MetricsData.filter(m => {
                    const title = String(m.title ?? '').trim().toLowerCase();
                    return title === 'delivered defect density_month' || title === 'delivered defect density' || title === 'delivered defect density (post production)' || title === 'post delivery defects';
                }).map(i => ({ USL: i.USL, LSL: i.LSL }));
                const deliveredLsl = deliveredDefectBounds[0]?.LSL ?? 0;
                const deliveredUsl = deliveredDefectBounds[0]?.USL ?? 0;
                const postDeliveryDefectsByMonth = new Map<string, number>();
                const postDeliveryComplexityByMonth = new Map<string, number>();
                const postDeliveryComplexityTracker = new Map<string, Set<string>>();
                (defectsData || []).forEach(defect => {
                    const phase = String(defect?.DefectDetectedPhase ?? '').trim().toLowerCase();
                    if (phase !== 'post production') {
                        return;
                    }
                    const defectDate = normalizeToLocalDateOnly(defect?.DefectDetectedOn ?? defect?.DefectClosureDate);
                    if (!defectDate) {
                        return;
                    }
                    const monthKey = `${defectDate.getFullYear()}-${defectDate.getMonth() + 1}`;
                    const requirementRaw = String(defect?.Requirement ?? '').trim();
                    if (!requirementRaw) {
                        return;
                    }
                    const resolvedRequirementKey = workItemsForDefectDensity.has(requirementRaw)
                        ? requirementRaw
                        : workItemLookup.get(requirementRaw.toLowerCase());
                    if (!resolvedRequirementKey) {
                        return;
                    }
                    postDeliveryDefectsByMonth.set(monthKey, (postDeliveryDefectsByMonth.get(monthKey) ?? 0) + 1);
                    const complexity = complexityByWorkItem.get(resolvedRequirementKey) ?? 0;
                    let tracker = postDeliveryComplexityTracker.get(monthKey);
                    if (!tracker) {
                        tracker = new Set<string>();
                        postDeliveryComplexityTracker.set(monthKey, tracker);
                    }
                    if (!tracker.has(resolvedRequirementKey)) {
                        tracker.add(resolvedRequirementKey);
                        postDeliveryComplexityByMonth.set(monthKey, (postDeliveryComplexityByMonth.get(monthKey) ?? 0) + complexity);
                    }
                });
                const combinedPostDeliveryKeys = new Set<string>();
                Object.keys(groupedByMonth).forEach(monthKey => combinedPostDeliveryKeys.add(monthKey));
                postDeliveryDefectsByMonth.forEach((_, monthKey) => combinedPostDeliveryKeys.add(monthKey));
                postDeliveryComplexityByMonth.forEach((_, monthKey) => combinedPostDeliveryKeys.add(monthKey));
                const postDeliveryDefectDensityTrendData: any[] = [['Month', 'Post Delivery Defect Density', { role: 'annotation', type: 'string' }, 'LSL', 'USL']];
                const postDeliveryMonthKeys: string[] = [];
                combinedPostDeliveryKeys.forEach(key => postDeliveryMonthKeys.push(key));
                postDeliveryMonthKeys.sort((a: string, b: string) => a.localeCompare(b));
                postDeliveryMonthKeys.forEach((monthKey: string) => {
                    const defectsForMonth = postDeliveryDefectsByMonth.get(monthKey) ?? 0;
                    const complexityForMonth = postDeliveryComplexityByMonth.get(monthKey) ?? 0;
                    const fallbackSize = groupedByMonth[monthKey]?.size ?? 0;
                    const denominator = complexityForMonth > 0 ? complexityForMonth : fallbackSize;
                    const density = denominator > 0 ? defectsForMonth / denominator : 0;
                    const roundedDensity = parseFloat(density.toFixed(3));
                    postDeliveryDefectDensityTrendData.push([
                        monthKey,
                        roundedDensity,
                        roundedDensity.toFixed(3),
                        deliveredLsl ?? 0,
                        deliveredUsl ?? 0
                    ]);
                });
                if (postDeliveryDefectDensityTrendData.length === 1) {
                    postDeliveryDefectDensityTrendData.push([
                        'No Data',
                        0,
                        '0.000',
                        deliveredLsl ?? 0,
                        deliveredUsl ?? 0
                    ]);
                }
                console.log('postDeliveryDefectDensityTrendData:', postDeliveryDefectDensityTrendData);
                setPostDeliveryDefectDensityChartData(postDeliveryDefectDensityTrendData);
                // Testing-phase defect density (system/integration/regression/acceptance)
                const testingPhases = new Set<string>(['system testing', 'integration testing', 'regression testing', 'acceptance testing'].map(s => s.toLowerCase()));
                const testingDefectBounds = MetricsData.filter(m => {
                    const title = String(m.title ?? '').trim().toLowerCase();
                    return title.indexOf('defect density') !== -1 || title.indexOf('testing defect') !== -1;
                }).map(i => ({ USL: i.USL, LSL: i.LSL }));
                const testingLsl = testingDefectBounds[0]?.LSL ?? 0;
                const testingUsl = testingDefectBounds[0]?.USL ?? 0;
                const testingDefectsByMonth = new Map<string, number>();
                const testingComplexityByMonth = new Map<string, number>();
                const testingComplexityTracker = new Map<string, Set<string>>();
                (defectsData || []).forEach(defect => {
                    const phase = String(defect?.DefectDetectedPhase ?? '').trim().toLowerCase();
                    if (!testingPhases.has(phase)) {
                        return;
                    }
                    const defectDate = normalizeToLocalDateOnly(defect?.DefectDetectedOn ?? defect?.DefectClosureDate);
                    if (!defectDate) {
                        return;
                    }
                    const monthKey = `${defectDate.getFullYear()}-${defectDate.getMonth() + 1}`;
                    const requirementRaw = String(defect?.Requirement).trim();
                    if (!requirementRaw) {
                        return;
                    }
                    const resolvedRequirementKey = workItemsForDefectDensity.has(requirementRaw)
                        ? requirementRaw
                        : workItemLookup.get(requirementRaw.toLowerCase());
                    if (!resolvedRequirementKey) {
                        return;
                    }
                    testingDefectsByMonth.set(monthKey, (testingDefectsByMonth.get(monthKey) ?? 0) + 1);
                    const complexity = complexityByWorkItem.get(resolvedRequirementKey) ?? 0;
                    let tracker = testingComplexityTracker.get(monthKey);
                    if (!tracker) {
                        tracker = new Set<string>();
                        testingComplexityTracker.set(monthKey, tracker);
                    }
                    if (!tracker.has(resolvedRequirementKey)) {
                        tracker.add(resolvedRequirementKey);
                        testingComplexityByMonth.set(monthKey, (testingComplexityByMonth.get(monthKey) ?? 0) + complexity);
                    }
                });
                const combinedTestingKeys = new Set<string>();
                Object.keys(groupedByMonth).forEach(monthKey => combinedTestingKeys.add(monthKey));
                testingDefectsByMonth.forEach((_, monthKey) => combinedTestingKeys.add(monthKey));
                testingComplexityByMonth.forEach((_, monthKey) => combinedTestingKeys.add(monthKey));
                const testingPhaseDefectDensityTrendData: any[] = [['Month', 'Testing Phase Defect Density', { role: 'annotation', type: 'string' }, 'LSL', 'USL']];
                const testingMonthKeys: string[] = [];
                combinedTestingKeys.forEach(key => testingMonthKeys.push(key));
                testingMonthKeys.sort((a: string, b: string) => a.localeCompare(b));
                testingMonthKeys.forEach((monthKey: string) => {
                    const defectsForMonth = testingDefectsByMonth.get(monthKey) ?? 0;
                    const complexityForMonth = testingComplexityByMonth.get(monthKey) ?? 0;
                    const fallbackSize = groupedByMonth[monthKey]?.size ?? 0;
                    const denominator = complexityForMonth > 0 ? complexityForMonth : fallbackSize;
                    const density = denominator > 0 ? defectsForMonth / denominator : 0;
                    const roundedDensity = parseFloat(density.toFixed(3));
                    testingPhaseDefectDensityTrendData.push([
                        monthKey,
                        roundedDensity,
                        roundedDensity.toFixed(3),
                        testingLsl ?? 0,
                        testingUsl ?? 0
                    ]);
                });
                if (testingPhaseDefectDensityTrendData.length === 1) {
                    testingPhaseDefectDensityTrendData.push([
                        'No Data',
                        0,
                        '0.000',
                        testingLsl ?? 0,
                        testingUsl ?? 0
                    ]);
                }
                console.log('testingPhaseDefectDensityTrendData:', testingPhaseDefectDensityTrendData);
                setTestingPhaseDefectDensityChartData(testingPhaseDefectDensityTrendData);
                


            if (selectedProjectType?.toLocaleLowerCase() == 'devm') {
                //Effort Variation Trend Data
                let USLLSLValuesEV = MetricsData.filter(m => m.title === 'Effort Variation_Month').map(i => {
                    return { USL: i.USL, LSL: i.LSL, Title: i.Title }
                });
                let effortVariationTrendData: any[] = []
                effortVariationTrendData.push(['Month', 'Effort Variation', { role: 'annotation', type: 'string' }, 'LSL', 'USL']);
                Object.keys(groupedByMonth).sort((a, b) => a.localeCompare(b)).forEach(monthKey => {
                    const effortVariation = ((groupedByMonth[monthKey].ActualEffort - groupedByMonth[monthKey].PlannedEffort) * 100) / groupedByMonth[monthKey].PlannedEffort;
                    effortVariationTrendData.push([monthKey, parseFloat(effortVariation.toFixed(2)), parseFloat(effortVariation.toFixed(2)), USLLSLValuesEV[0]?.LSL ?? 0, USLLSLValuesEV[0]?.USL ?? 0]);
                });
                setefforVatiationChartData(effortVariationTrendData);
                //Overall Productivity Trend Data
                let overallProductivityTrendData: any[] = []
                overallProductivityTrendData.push(['Month', 'Overall Productivity', { role: 'annotation', type: 'string' }, 'LSL', 'USL']);
                let USLLSLValuesOAP = MetricsData.filter(m => m.title === 'Overall Productivity_Month').map(i => {
                    return { USL: i.USL, LSL: i.LSL, Title: i.Title }
                });
                Object.keys(groupedByMonth).sort((a, b) => a.localeCompare(b)).forEach(monthKey => {
                    const overallProductivity = (groupedByMonth[monthKey].ActualEffort) / (groupedByMonth[monthKey].size);
                    overallProductivityTrendData.push([monthKey, parseFloat(overallProductivity.toFixed(2)), parseFloat(overallProductivity.toFixed(2)), USLLSLValuesOAP[0]?.LSL ?? 0, USLLSLValuesOAP[0]?.USL ?? 0]);
                });
                setOverallProductivityChartData(overallProductivityTrendData);

               

                //Requirement Analysis Effort Density Trend Data
                let requirementAnalysisEffortDensityTrendData: any[] = [];
                requirementAnalysisEffortDensityTrendData.push(['Month', 'Requirement Analysis Effort Density', { role: 'annotation', type: 'string' }]);
                let codingProductivityTrendData: any[] = [];
                codingProductivityTrendData.push(['Month', 'Coding Productivity', { role: 'annotation', type: 'string' }]);
                let CodeReviewEffortDensityChartData: any[] = [];
                CodeReviewEffortDensityChartData.push(['Month', 'Code Review Effort Density', { role: 'annotation', type: 'string' }]);
                let CodeReworkEffortDensity = 0.0;
                let UnitTestingEffortDensity = 0.0;
                let TestingExecutionEffortDensity = 0.0;
                let CodeReworkEffortDensityChartData: any[] = [];
                CodeReworkEffortDensityChartData.push(['Month', 'Code Rework Effort Density', { role: 'annotation', type: 'string' }]);
                let UnitTestingEffortDensityChartData: any[] = [];
                UnitTestingEffortDensityChartData.push(['Month', 'Unit Testing Effort Density', { role: 'annotation', type: 'string' }]);
                let TestingExecutionEffortDensityChartData: any[] = [];
                TestingExecutionEffortDensityChartData.push(['Month', 'Testing Execution Effort Density', { role: 'annotation', type: 'string' }]);
                // let USLLSLValuesRAED = MetricsData.filter(m => m.title === 'Requirement Analysis Effort Density_Month').map(i => {
                //     return { USL: i.USL, LSL: i.LSL, Title: i.Title }
                // });      
                Object.keys(groupedByMonth).sort((a, b) => a.localeCompare(b)).forEach(monthKey => {
                    let reqAnalysisEffortForWorkItem = 0.0;
                    let codingProductivity = 0.0;
                    let CodeReviewEffortDensity = 0.0;
                    WorkLogItemWithPlannedandActualEfforts.forEach(wlItem => {
                        const actualEndDate = normalizeToLocalDateOnly(wlItem.ActualEndDate);
                        if (actualEndDate) {
                            const wlMonthKey = `${actualEndDate.getFullYear()}-${actualEndDate.getMonth() + 1}`;
                            if (wlMonthKey === monthKey) {
                                const taskType = (typeof wlItem.TaskType === 'string' ? wlItem.TaskType : '').toLowerCase();
                                //  closedTaskByWorkItem.forEach(task => {
                                if (taskType === 'requirements analysis' && wlItem.AdjustedComplexityPoint > 0) {
                                    reqAnalysisEffortForWorkItem += (wlItem.ActualEffort) / (wlItem.AdjustedComplexityPoint);
                                }
                                if (taskType === 'coding' && wlItem.AdjustedComplexityPoint > 0) {
                                    codingProductivity += (wlItem.ActualEffort) / (wlItem.AdjustedComplexityPoint);
                                }
                                if (taskType === 'code review' && wlItem.AdjustedComplexityPoint > 0) {
                                    CodeReviewEffortDensity += (wlItem.ActualEffort / wlItem.AdjustedComplexityPoint);
                                }
                                if (taskType === 'rework after code review' && wlItem.AdjustedComplexityPoint > 0) {
                                    CodeReworkEffortDensity += (wlItem.ActualEffort / wlItem.AdjustedComplexityPoint);
                                }
                                if (taskType === 'unit testing' && wlItem.AdjustedComplexityPoint > 0) {
                                    UnitTestingEffortDensity += (wlItem.ActualEffort / wlItem.AdjustedComplexityPoint);
                                }
                                if ((taskType === 'system testing' || taskType === 'integration testing') && wlItem.AdjustedComplexityPoint > 0) {
                                    TestingExecutionEffortDensity += (wlItem.ActualEffort / wlItem.AdjustedComplexityPoint);
                                }
                                // });
                                // totalReqAnalysisEffort += reqAnalysisEffortForWorkItem;
                                // totalSize += wlItem.size;
                            }
                        }
                    });
                    requirementAnalysisEffortDensityTrendData.push([monthKey, parseFloat(reqAnalysisEffortForWorkItem.toFixed(2)), parseFloat(reqAnalysisEffortForWorkItem.toFixed(2))]);
                    codingProductivityTrendData.push([monthKey, parseFloat(codingProductivity.toFixed(2)), parseFloat(codingProductivity.toFixed(2))]);
                    CodeReviewEffortDensityChartData.push([monthKey, parseFloat(CodeReviewEffortDensity.toFixed(2)), parseFloat(CodeReviewEffortDensity.toFixed(2))]);
                    CodeReworkEffortDensityChartData.push([monthKey, parseFloat(CodeReworkEffortDensity.toFixed(2)), parseFloat(CodeReworkEffortDensity.toFixed(2))]);
                    UnitTestingEffortDensityChartData.push([monthKey, parseFloat(UnitTestingEffortDensity.toFixed(2)), parseFloat(UnitTestingEffortDensity.toFixed(2))]);
                    TestingExecutionEffortDensityChartData.push([monthKey, parseFloat(TestingExecutionEffortDensity.toFixed(2)), parseFloat(TestingExecutionEffortDensity.toFixed(2))]);
                });
                setRequirementAnalysisEffortDensityChartData(requirementAnalysisEffortDensityTrendData);
                setCodingProductivityChartData(codingProductivityTrendData);
                setCodeReviewEffortDensityChartData(CodeReviewEffortDensityChartData);
                setCodeReworkEffortDensityChartData(CodeReworkEffortDensityChartData);
                setUnitTestingEffortDensityChartData(UnitTestingEffortDensityChartData);
                setTestingExecutionEffortDensityChartData(TestingExecutionEffortDensityChartData);

                //Scheduled Variation Trend Data
                let scheduledVatiationTrendData: any[] = []
                scheduledVatiationTrendData.push(['Month', 'Scheduled Variation', { role: 'annotation', type: 'string' }, 'LSL', 'USL']);
                let USLLSLValuesSVTrend = MetricsData.filter(m => m.title === 'Schedule Variation_Month').map(i => {
                    return { USL: i.USL, LSL: i.LSL, Title: i.Title }
                });
                Object.keys(groupedByMonth).sort((a, b) => a.localeCompare(b)).forEach(monthKey => {
                    const scheduledVariation = (groupedByMonth[monthKey].plannedDuration !== 0) ? ((groupedByMonth[monthKey].ActualEffort - groupedByMonth[monthKey].PlannedEffort) * 100) / groupedByMonth[monthKey].plannedDuration : 0;
                    scheduledVatiationTrendData.push([monthKey, parseFloat(scheduledVariation.toFixed(2)), parseFloat(scheduledVariation.toFixed(2)), USLLSLValuesSVTrend[0]?.LSL ?? 0, USLLSLValuesSVTrend[0]?.USL ?? 0]);
                });
                setscheduledVatiationChartData(scheduledVatiationTrendData);

            }
            if (selectedProjectType?.toLocaleLowerCase() === 'dev') {
                //Create a chart for Effort Variation
                //if (WorkLogItemWithPlannedandActualEfforts.length > 0) {
                let EfforVatiationChartDataDEV: any[] = []
                EfforVatiationChartDataDEV.push(['WorkItemNo', 'Effort Variation', 'LSL', 'USL']);
                let USLLSLValuesDEV = MetricsData.filter(m => m.title === 'Effort Variation').map(i => {
                    return { USL: i.USL, LSL: i.LSL, Title: i.Title }
                });
                WorkLogItemWithPlannedandActualEfforts.forEach(ev => {
                    //new Date(ev.ActualEndDate).toLocaleString('en-US', { month: 'short' })
                    EfforVatiationChartDataDEV.push([ev.WorkItemNo, ev.EffortVariation, USLLSLValuesDEV[0]?.LSL ?? 0, USLLSLValuesDEV[0]?.USL ?? 0])
                });
                setefforVatiationChartData(EfforVatiationChartDataDEV);

                //Create chart for scheduled Variation
                let ScheduledVatiationChartDataDEV: any[] = []
                ScheduledVatiationChartDataDEV.push(['WorkItemNo', 'Schedule Variation', 'LSL', 'USL']);
                let USLLSLValuesSVDEV = MetricsData.filter(m => m.title === 'Schedule Variation').map(i => {
                    return { USL: i.USL, LSL: i.LSL, Title: i.Title }
                });
                WorkLogItemWithPlannedandActualEfforts.forEach(sv => {
                    //new Date(ev.ActualEndDate).toLocaleString('en-US', { month: 'short' })
                    ScheduledVatiationChartDataDEV.push([sv.WorkItemNo, sv.ScheduledVariation, USLLSLValuesSVDEV[0]?.LSL ?? 0, USLLSLValuesSVDEV[0]?.USL ?? 0])
                });
                setscheduledVatiationChartData(ScheduledVatiationChartDataDEV);

                //Create a chart for Overall productivity

                let OverallProductivityChartDataDEV: any[] = []
                OverallProductivityChartDataDEV.push(['WorkItemNo', 'Overall Productivity', 'LSL', 'USL']);
                let USLLSLValuesOPDEV = MetricsData.filter(m => m.title === 'Overall Productivity').map(i => {
                    return { USL: i.USL, LSL: i.LSL, Title: i.Title }
                });
                WorkLogItemWithPlannedandActualEfforts.forEach(op => {
                    //new Date(ev.ActualEndDate).toLocaleString('en-US', { month: 'short' })
                    OverallProductivityChartDataDEV.push([op.WorkItemNo, op.OverAllProductivity, USLLSLValuesOPDEV[0]?.LSL ?? 0, USLLSLValuesOPDEV[0]?.USL ?? 0])
                });
                console.log('overall productivity', OverallProductivityChartDataDEV)
                setOverallProductivityChartData(OverallProductivityChartDataDEV);

                let RequirementAnalysisEffortDensityChartDataDEV: any[] = []
                RequirementAnalysisEffortDensityChartDataDEV.push(['WorkItemNo', 'Requirement Analysis Effort Density']);
                let codingProductivityChartDataDEV: any[] = []
                codingProductivityChartDataDEV.push(['WorkItemNo', 'Coding Productivity']);
                let codeReviewEffortDensityChartDataDEV: any[] = []
                codeReviewEffortDensityChartDataDEV.push(['WorkItemNo', 'Code Review Effort Density']);
                let codeReworkEffortDensityChartDataDEV: any[] = []
                codeReworkEffortDensityChartDataDEV.push(['WorkItemNo', 'Code Rework Effort Density']);
                let unitTestingEffortDensityChartDataDEV: any[] = [];
                unitTestingEffortDensityChartDataDEV.push(['WorkItemNo', 'Unit Testing Effort Density']);
                let testingExecutionEffortDensityChartDataDEV: any[] = [];
                testingExecutionEffortDensityChartDataDEV.push(['WorkItemNo', 'Testing Execution Effort Density']);
                // let USLLSLValuesRAEDDEV = MetricsData.filter(m => m.title === 'Requirement Analysis Effort Density').map(i => {
                //     return { USL: i.USL, LSL: i.LSL, Title: i.Title }
                // });
                WorkLogItemWithPlannedandActualEfforts.forEach(raed => {
                    //new Date(ev.ActualEndDate).toLocaleString('en-US', { month: 'short' })
                    RequirementAnalysisEffortDensityChartDataDEV.push([raed.WorkItemNo, raed.ReqAnalyisEffortDensity ?? 0])

                });
                setRequirementAnalysisEffortDensityChartData(RequirementAnalysisEffortDensityChartDataDEV);
                WorkLogItemWithPlannedandActualEfforts.forEach(cp => {
                    codingProductivityChartDataDEV.push([cp.WorkItemNo, cp.codingProductivity ?? 0])
                    //new Date(ev.ActualEndDate).toLocaleString('en-US', { month: 'short' })
                });
                console.log('coding productivity', codingProductivityChartDataDEV)
                setCodingProductivityChartData(codingProductivityChartDataDEV);
                WorkLogItemWithPlannedandActualEfforts.forEach(cR => {
                    codeReviewEffortDensityChartDataDEV.push([cR.WorkItemNo, cR.CodeReviewEffortDensity ?? 0])
                    codeReworkEffortDensityChartDataDEV.push([cR.WorkItemNo, cR.CodeReworkEffortDensity ?? 0])
                    unitTestingEffortDensityChartDataDEV.push([cR.WorkItemNo, cR.UnitTestingEffortDensity ?? 0])
                    testingExecutionEffortDensityChartDataDEV.push([cR.WorkItemNo, cR.TestingExecutionEffortDensity ?? 0])
                    //new Date(ev.ActualEndDate).toLocaleString('en-US', { month: 'short' })
                });
                console.log('coding Review EffortDensity', codeReviewEffortDensityChartDataDEV)
                setCodeReviewEffortDensityChartData(codeReviewEffortDensityChartDataDEV);
                setCodeReworkEffortDensityChartData(codeReworkEffortDensityChartDataDEV);
                setUnitTestingEffortDensityChartData(unitTestingEffortDensityChartDataDEV);
                setTestingExecutionEffortDensityChartData(testingExecutionEffortDensityChartDataDEV);

            }







            // setOverallProductivityChartData(OverallProductivityChartData);
            console.log('Mean of Effort Variation', meanEffortVariation);
            console.log('Mean of Effort Variation', MeanEffortVariation);
            console.log('Mean of Effort Variation', MeanSV);
            console.log('Mean OverAllProductivity', MeanOverallProductivity)
            console.log('Standard Deviation of Effort Variation', StandardDeviationOfEV);
            console.log('Standard Deviation of Scheduled Variation', StandardDeviationSV);
            setFilteredWorkLogData(WorkLogItemWithPlannedandActualEfforts);
            console.log('Filtered WorkLogData result:', WorkLogItemWithPlannedandActualEfforts);
            console.log('Filtered WorkLogData:', filteredWorkLogData);

        }
    }, [TaskManagementData, WorkLogData, MetricsData, selectedMonth, defectsData, CodeReviewDefectsData, AMSEffortLogData, PPOApproversData, ManagementEffortLogData, selectedProjectType]);
    {/* -------------------------------- DEVM Dashboard Data calculations Ended ---------------------- */ }

    //-------------------------------- Testing Defects Data Load Started ----------------------  */}
    const testingDefectDataForDashboard = async (WorkItemActualEffort: any[]) => {
        try {
            const monthMatches = (raw: any) => {
                if (!raw) return false;
                const d = raw instanceof Date ? raw : new Date(raw);
                return !isNaN(d.getTime()) && d.getMonth().toString() === String(selectedMonth);
            };
            const findEffortForReq = (req: any) => {
                const id = String(req ?? '').trim();
                if (!id) return 0;
                const match = (WorkItemActualEffort || []).filter(w =>
                    String(w.WorkItemNo) === id && monthMatches(w.ActualEndDate)
                )[0];
                return match ? Number(match.ActualEffort) || 0 : 0;
            };

            const codeReviewDefs = (CodeReviewDefectsData ?? []).filter(cr =>
                String(cr.DefectType ?? '').trim().toLowerCase() === 'code review' &&
                (monthMatches(cr.IdentifiedDate ?? null) || monthMatches(cr.ReviewCompletionDate ?? null))
            );
            const systemTypes = new Set(['system testing', 'integration testing', 'regression testing']);
            const systemDefs = (defectsData ?? []).filter(d =>
                systemTypes.has(String(d.TestingType ?? '').trim().toLowerCase()) &&
                monthMatches(d.DefectDetectedOn)
            );
            const unitDefs = (defectsData ?? []).filter(d =>
                String(d.TestingType ?? '').trim().toLowerCase() === 'unit testing' &&
                monthMatches(d.DefectDetectedOn)
            );

            const items = [
                {
                    key: 'codeReview',
                    Category: 'Code Review',
                    DefectCount: codeReviewDefs.length,
                    TotalEffort: parseFloat(codeReviewDefs.reduce((sum, cr) => sum + findEffortForReq(cr.RequirementID ?? cr.Requirement), 0).toFixed(2))
                },
                {
                    key: 'systemGroup',
                    Category: 'System/Integration/Regression',
                    DefectCount: systemDefs.length,
                    TotalEffort: parseFloat(systemDefs.reduce((sum, d) => sum + findEffortForReq(d.Requirement ?? d.Requirement), 0).toFixed(2))
                },
                {
                    key: 'unitTesting',
                    Category: 'Unit Testing',
                    DefectCount: unitDefs.length,
                    TotalEffort: parseFloat(unitDefs.reduce((sum, d) => sum + findEffortForReq(d.Requirement ?? d.Requirement), 0).toFixed(2))
                }
            ];
            const columns: IColumn[] = [
                { key: 'colCategory', name: 'Category', fieldName: 'Category', minWidth: 180 },
                { key: 'colCount', name: 'Defect Count', fieldName: 'DefectCount', minWidth: 100 },
                { key: 'colEffort', name: 'Total Effort', fieldName: 'TotalEffort', minWidth: 100 }
            ];

            setDefectsDetailsListItems(items);
            setDefectsDetailsListColumns(columns);
        } catch (err) {
            console.error('Defects summary calculation failed', err);
        }
    }



    //Testing Defects Data
    React.useEffect(() => {
        if (context) {
            loadDefectsData().catch(() => { });
        }
    }, [context]);

    React.useEffect(() => {
        if (context) {
            loadAMSTicketsData().catch(() => { });
        }
    }, [context]);
    const loadAMSTicketsData = async () => {
        if (!context) return;
        const genericServiceInstance: IGenericService = new GenericService(undefined, context);
        genericServiceInstance.init(undefined, context);
        const AMSTicketsRepo: IAMSEffortLogRepository = new AMSEffortLogRepository(genericServiceInstance);
        AMSTicketsRepo.setService(genericServiceInstance);
        let AMSTicketsValues = await getIAMSEffortLogValues(false, context, selectedProjectType);
        const mapped = AMSTicketsValues.map(m => ({
            Title: m.Title,
            ActualEfforts: m.ActualEfforts,
            ActualStartDate: m.ActualStartDate,
            TaskType: m.TaskType,
        }));
        setAMSEffortLogData(mapped);
        console.log('AMSEffortLogData:', AMSEffortLogData);
    }

    const loadDefectsData = async () => {
        if (!context) return;
        const genericServiceInstance: IGenericService = new GenericService(undefined, context);
        genericServiceInstance.init(undefined, context);
        const DefectsRepo: ITestingDefectsRepository = new TestingDefectsRepository(genericServiceInstance);
        DefectsRepo.setService(genericServiceInstance);
        let DefectsValues = await getTestingDefectsValues(false, context, selectedProjectType);
        if (!DefectsValues || DefectsValues.length === 0) return;
        const mapped = DefectsValues.map(m => ({
            Requirement: m.Requirement,
            TestScenarioID: m.TestScenarioID,
            TestCaseID: m.TestCaseID,
            DefectDescription: m.DefectDescription,
            TestingType: m.TestingType,
            DefectDetectedOn: m.DefectDetectedOn,
            DefectDetectedBy: m.DefectDetectedBy,
            DefectStatus: m.DefectStatus,
            // DefectType: m.DefectType,
            // DefectClassification: m.DefectClassification,
            DefectOriginPhase: m.DefectOriginPhase,
            DefectDetectedPhase: m.DefectDetectedPhase,
            Severity: m.Severity,
            Priority: m.Priority,
            RootCause: m.RootCause,
            DefectFixedBy: m.DefectFixedBy,
            DefectClosureDate: m.DefectClosureDate,
            Remarks: m.Remarks

        }))
        setDefectsData(mapped);
        console.log('DefectsData:', defectsData);
    };
    React.useEffect(() => {
        if (context) {
            loadCodeReviewDefectsData().catch(() => { });
        }
    }, [context]);

    const loadCodeReviewDefectsData = async () => {
        if (!context) return;
        const genericServiceInstance: IGenericService = new GenericService(undefined, context);
        genericServiceInstance.init(undefined, context);
        const DefectsRepo: ICodeReviewDefectsRepository = new CodeReviewDefectRepository(genericServiceInstance);
        DefectsRepo.setService(genericServiceInstance);
        let DefectsValues = await getCodeReviewDefectsValues(false, context, selectedProjectType);
        if (!DefectsValues || DefectsValues.length === 0) return;
        const mapped = DefectsValues.map(it => ({
            RequirementID: it.RequirementID,

            CodeFileClassName: it?.CodeFileClassName,
            CodeFileClassSize: it?.CodeFileClassSize,
            CodeFileAuthor: it?.CodeFileAuthor,
            ReviewerName: it?.ReviewerName,
            ReviewIterationNumber: it?.ReviewIterationNumber,
            IdentifiedDate: it?.IdentifiedDate,
            ReviewCompletionDate: it?.ReviewCompletionDate,
            DefectDescription: it?.DefectDescription,
            CodeReviewChecklist: it?.CodeReviewChecklist,
            ReviewResults: it?.ReviewResults,
            DefectStatus: it?.DefectStatus,
            DefectType: it?.DefectType,
            DefectClassification: it?.DefectClassification,
            DefectOriginPhase: it?.DefectOriginPhase,
            impactedComponents: it?.impactedComponents,
            CorrectionCorrectiveAction: it?.CorrectionCorrectiveAction,
            PlannedClosureDate: it?.PlannedClosureDate,
            ActualClosureDate: it?.ActualClosureDate,
            LocationOfDefectInCode: it?.LocationOfDefectInCode,
            Severity: it?.Severity,
            Remarks: it?.Remarks

        }))
        setCodeReviewDefectsData(mapped);
        console.log('DefectsData:', CodeReviewDefectsData);
    };


    //customer statisfaction index trends data
    React.useEffect(() => {
        if (csiData.length > 0 && MetricsData.length > 0 && context) {
            csiTrendData();
        }


    }, [csiData, MetricsData, context]);


    function parseDDMMYYYY(input: string): Date {
        if (typeof input !== 'string') return new Date(NaN);

        const s = input.trim();

        // Remove time part if present (e.g., "06/01/2026 12:00 AM")
        const datePart = s.split(' ')[0];

        // Accept separators: / - . or mixed
        const parts = datePart.split(/[\/\-.]/);
        if (parts.length !== 3) return new Date(NaN);

        let [ddStr, mmStr, yyyyStr] = parts.map(p => p.trim());

        // Handle 2-digit year by assuming 2000–2099 (adjust to your business rule)
        if (yyyyStr.length === 2) yyyyStr = '20' + yyyyStr;

        const dd = Number(ddStr);
        const mm = Number(mmStr);
        const yyyy = Number(yyyyStr);

        // Basic numeric checks
        if (!Number(dd) || !Number(mm) || !Number(yyyy)) return new Date(NaN);
        if (mm < 1 || mm > 12) return new Date(NaN);

        // Days in month (handles leap years)
        const daysInMonth = new Date(yyyy, mm, 0).getDate(); // using local month to compute count
        if (dd < 1 || dd > daysInMonth) return new Date(NaN);

        // Construct date in UTC to avoid timezone shift
        return new Date(Date.UTC(yyyy, mm - 1, dd));
    }



    const csiTrendData = (() => {
        if (!csiData || csiData.length === 0) return [];
        let USLLSLValues: any[] = [];
        if (MetricsData.length > 0) {
            USLLSLValues = MetricsData.filter(m => m.title.toLowerCase() === 'customer satisfaction index').map(i => {
                return { USL: i.USL, LSL: i.LSL, Title: i.Title }
            });
        }
        let csiTrendataForChart: any[] = [['Date', 'CSI', { role: 'annotation', type: 'string' }, 'Lower Limit', 'Upper Limit']];
        if (USLLSLValues.length > 0) {


            const rows = csiData
                .slice()
                .sort((a: any, b: any) => new Date(a.CSATAquiredDate).getTime() - new Date(b.CSATAquiredDate).getTime())
                .map((it: any) => {

                    // const d = parseDDMMYYYY(it.CSATAquiredDate);
                    //  console.log('date',d)
                    // const label = d.toLocaleDateString('en-US');
                    const value = Number(it.Title);
                    csiTrendataForChart.push([new Date(it.CSATAquiredDate).toLocaleDateString('en-US'), value, value, Number(USLLSLValues[0]?.LSL ?? 0), Number(USLLSLValues[0]?.USL ?? 0)]);
                });

        }

        setCSITrendData(csiTrendataForChart);
        console.log('csiTrendataForChart', csiTrendataForChart)

    });

    const options = {
        legend: { position: 'none' },
        seriesType: 'line',
        series: {

            0: { color: '#28a745', lineWidth: 2, pointsVisible: true, pointSize: 6, pointShape: 'circle' }, // Customer Rating
            1: { color: '#dc3545', lineWidth: 2, lineDashStyle: [4, 4], pointsVisible: true, pointSize: 6, pointShape: 'triangle' }, // Target
        },
        colors: ['#28a745', '#dc3545'], // keep in sync with series if you use colors
        annotations: {
            alwaysOutside: true,
            textStyle: { fontSize: 11, bold: true, color: '#333' },
            stem: { color: 'transparent', length: 0 }, // hide stems if overlaying points
        },
        // Optional: reduce hover-only info if labels are always visible
        // tooltip: { trigger: 'selection' }, // or 'none' to turn off
        // Make points more prominent against lines:
        dataOpacity: 0.9,
        // Make gridlines subtle:
        hAxis: { textStyle: { color: '#555' } },
        vAxis: { textStyle: { color: '#555' }, gridlines: { color: '#eee' } },
    };

    React.useEffect(() => {
        if (context) {
            loadFacilitationData().catch(() => { });
        }
    }, [context]);
    React.useEffect(() => {
        facilitationChartData();
    }, [facilitationReportData]);

    const loadFacilitationData = async () => {
        if (!context) return;
        const genericServiceInstance: IGenericService = new GenericService(undefined, context);
        genericServiceInstance.init(undefined, context);
        const FacilitationRepo: IFacilitationReportRepository = new FacilitationReportRepository(genericServiceInstance);
        FacilitationRepo.setService(genericServiceInstance);
        let FacilitationValues = await getFacilitationValues(false, context, selectedProjectType);
        const mapped = FacilitationValues.map(m => ({
            ID: m.ID,
            Category: m.Category,
            Finding: m.Finding,
            FindingDate: m.FindingDate,
            ClosureDate: m.ClosureDate,
            Status: m.Status,
        }));
        setFacilitationReportData(mapped);
        console.log('FacilitationReportData:', facilitationReportData);
    }

    const daysFromNow = (findingDate: any): number => {
        const target = new Date(findingDate);
        const now = new Date();
        const msPerDay = 24 * 60 * 60 * 1000;

        // Difference in milliseconds
        const diffMs = target.getTime() - now.getTime();

        // Whole days (discard partial day)
        return Math.floor(diffMs / msPerDay);
    };



    const facilitationChartData = (() => {
        if (!facilitationReportData || facilitationReportData.length === 0) return [];
        let facilitationDataForChart: any[] = [['Category', 'Count']];
        const categoryCountMap = new Map<string, number>();
        facilitationReportData.forEach(item => {
            const category = item.Category || 'Uncategorized';
            categoryCountMap.set(category, (categoryCountMap.get(category) || 0) + 1);
        });
        categoryCountMap.forEach((count, category) => {
            facilitationDataForChart.push([category, count]);
        });
        let data = facilitationReportData.filter(i => i.Category === 'FC Finding')
        let AgingData = [];
        AgingData.push(['Aging', 'Findings']);
        //daysFromNow(x.FindingDate)
        AgingData.push(['>=7 days', data.filter(x => daysFromNow(x.FindingDate) >= 7).length])
        AgingData.push(['<=30 days', data.filter(x => daysFromNow(x.FindingDate) <= 30).length])
        AgingData.push(['>30 days', data.filter(x => daysFromNow(x.FindingDate) > 30).length])
        //     ['Aging', 'Findings'],
        //     ['>=7 days', 10],
        //     ['<=30 days', 2],
        //     ['>30 days', 5],


        setFacilitationReportDataForChart(facilitationDataForChart);
        setagingFindingsData(AgingData);
        console.log(AgingData);
    });
    React.useEffect(() => {
        if (context && Array.isArray(PPOApproversData) && PPOApproversData.length > 0 && ManagementEffortLogData.length > 0 && facilitationReportData.length > 0) {
            let pci = buildPCIChartData();
            setpciChartData(pci);
        }
    }, [context, PPOApproversData, ManagementEffortLogData, facilitationReportData]);

    // Build PCI chart data:
    function buildPCIChartData(): any[] {
        try {
            const rows: any[] = [['Date', 'PCI', { role: 'annotation', type: 'string' }]];

            const siteUrl = context?.pageContext?.web?.absoluteUrl || '';

            // gather reviewer identifiers for current site (be tolerant of field names)
            const reviewerValues = new Set<string>();
            if (Array.isArray(PPOApproversData)) {
                PPOApproversData.forEach(p => {
                    if (!p) return;
                    if (String(p.SiteURL || '') !== String(siteUrl)) return;
                    const candidates = [p.ReviewerName, p.Reviewer, p.ReviewerEmail, p.Email, p.Title];
                    candidates.forEach(c => {
                        if (c) reviewerValues.add(String(c).trim().toLowerCase());
                    });
                });
            }

            // filter management effort entries where reviewer matches and activity is audit(s)
            const managementFiltered = (ManagementEffortLogData || []).filter(m => {
                if (!m) return false;
                const activity = String(m.ManagementTaskActivity || '').trim().toLowerCase();
                if (activity.indexOf('audits') === -1) return false;
                const updatedBy = String(m.UpdatedBy || '').trim().toLowerCase();
                if (!updatedBy) return false;
                if (reviewerValues.size === 0) return false;
                return reviewerValues.has(updatedBy);
            });

            // build map of findingDate -> { findings: n, matchedAudits: m, nc, observation, fcFinding }
            const bucket = new Map<string, any>();

            const normalizeKey = (raw: any) => {
                if (!raw) return null;
                const d = raw instanceof Date ? raw : new Date(raw);
                if (isNaN(d.getTime())) return null;
                return `${d.getFullYear()}-${('0' + (d.getMonth() + 1)).slice(-2)}-${('0' + d.getDate()).slice(-2)}`;
            };

            const formatLabel = (key: string) => {
                const parts = key.split('-');
                if (parts.length !== 3) return key;
                const yyyy = Number(parts[0]);
                const mm = Number(parts[1]);
                const dd = Number(parts[2]);
                const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
                const yy = String(yyyy).slice(-2);
                return `${('0' + dd).slice(-2)}-${months[mm - 1]}-${yy}`;
            };

            // count findings grouped by FindingDate (use facilitationReportData's FindingDate)
            // track open counts for NC, Observation and FC Finding per date
            (facilitationReportData || []).forEach(f => {
                const key = normalizeKey(f?.FindingDate || f?.FindingDate);
                if (!key) return;
                const cur = bucket.get(key) || { findings: 0, matched: 0, nc: 0, observation: 0, fcFinding: 0 } as any;
                cur.findings += 1;
                const status = String(f?.Status || '').trim().toLowerCase();
                const cat = String(f?.Category || '').trim().toLowerCase();
                if (status === 'open') {
                    if (cat.indexOf('nc') !== -1) cur.nc += 1;
                    else if (cat.indexOf('observation') !== -1) cur.observation += 1;
                    else if (cat.indexOf('fc finding') !== -1 || (cat.indexOf('fc') !== -1 && cat.indexOf('finding') !== -1) || cat === 'fc') cur.fcFinding += 1;
                }
                bucket.set(key, cur);
            });

            // for each management entry, check if ActualStartDate matches any finding date and increment matched
            (managementFiltered || []).forEach(m => {
                const key = normalizeKey(m?.ActualStartDate);
                if (!key) return;
                const cur = bucket.get(key) || { findings: 0, matched: 0 };
                cur.matched += 1;
                bucket.set(key, cur);
            });

            // prepare sorted rows
            const keys: string[] = [];
            bucket.forEach((_, k) => keys.push(k));
            keys.sort();

            keys.forEach(k => {
                const val = bucket.get(k)!;
                // PCI: negative weighted sum of open counts per user request
                // PCIvalue == -(20*openNCCount + 10*openObservationCount + 10*openFCFindingCount)
                const openNC = Number(val.nc || 0);
                const openObs = Number(val.observation || 0);
                const openFC = Number(val.fcFinding || 0);
                console.log(`PCI for ${k}: NC=${openNC}, Obs=${openObs}, FC=${openFC}`);
                const pci = parseFloat((-((20 * openNC) + (10 * openObs) + (10 * openFC))).toFixed(2));
                rows.push([formatLabel(k), pci, String(pci)]);
            });

            // if no rows, return a small placeholder series to avoid empty chart errors
            if (rows.length === 1) {
                rows.push([formatLabel(new Date().toISOString().slice(0, 10)), 0, '0']);
            }

            return rows;
        } catch (err) {
            console.error('Failed to build PCI chart data', err);
            return [['Date', 'PCI', { role: 'annotation', type: 'string' }], [new Date().toLocaleDateString('en-US'), 0, '0']];
        }
    }
    React.useEffect(() => {
        fetchRCAItems();
    }, [context]);

    const fetchRCAItems = async () => {
        const genericServiceInstance: IGenericService = new GenericService(undefined, context);
        genericServiceInstance.init(undefined, context);
        const RCARepo: IRCARepository = new RCARepository(genericServiceInstance);
        RCARepo.setService(genericServiceInstance);
        const RAitems = await getRCAItems(true, context);
        setRCAItems(RAitems);
    }


    return (
        <div style={containerStyle}>
            <div style={contentStyle}>
                <header style={headerStyle}>
                    <Text variant="xLargePlus">Metrics Dashboard</Text>

                    <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
                        <Dropdown
                            selectedKey={selectedProjectType}
                            onChange={(_, value) => handleProjectTypeChange(value)}
                            options={ProjectType}
                            styles={{ root: { width: 180 } }}
                        />

                        <Dropdown
                            selectedKey={selectedMonth}
                            onChange={(_, option) => {
                                setselectedMonth(option?.key as string)
                            }
                            }
                            options={Months}
                            styles={{ root: { width: 180 } }}
                        />



                        {/*<PrimaryButton text="Export to XLSX" onClick={exportToCsv} />*/}
                    </Stack>
                </header>


                <section style={{ margin: '20px 0' }}>
                    <div style={kpiAreaStyle}>
                        {kpis.map(k => {
                            const pivotKey = k.pivotKey || k.title;
                            const isActive = selectedPivotKey === pivotKey;

                            return (
                                <div

                                    key={k.title}
                                    onClick={() => openDialogForPivot(pivotKey)}
                                    style={{
                                        ...kpiCardStyle(),
                                        cursor: 'pointer',
                                        border: isActive ? '2px solid #0078d4' : '1px solid rgba(16,24,40,0.04)',
                                    }}
                                    onMouseEnter={e => (e.currentTarget.style.transform = 'translateY(-4px)')}
                                    onMouseLeave={e => (e.currentTarget.style.transform = 'none')}
                                >
                                    <div>
                                        <Text variant="small" styles={{ root: { color: '#666' } }}>{k.title}</Text>
                                        <div style={{
                                            fontSize: 26,
                                            fontWeight: 700,
                                            marginTop: 8,
                                            marginBottom: 6,
                                            color: k.status === 'green' ? '#28a745' : '#dc3545',
                                        }} dangerouslySetInnerHTML={{ __html: k.value }} />
                                        <Text variant="small" styles={{ root: { color: '#888' } }}>Goal: {k.goal} </Text>
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                </section>
                <Dialog
                    hidden={!isDialogOpen}
                    onDismiss={closeDialog}
                    dialogContentProps={{
                        type: DialogType.largeHeader,
                        title: dialogPivotKey == 'Defect density' ? 'Defect Density' : dialogPivotKey || '',
                        closeButtonAriaLabel: 'Close dialog',
                        topButtonsProps: [
                            {
                                iconProps: { iconName: 'Cancel' },
                                ariaLabel: 'Close dialog',
                                onClick: closeDialog,
                            }
                        ]
                    }}
                    modalProps={{ isBlocking: false }}
                    minWidth={600}
                    maxWidth={900}
                >
                    <div style={{ minWidth: 480 }}>
                        {dialogPivotKey === 'Customer Satisfaction Index' && (
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Customer Satisfaction Index trend</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-rating`} chartKey="Customer Satisfaction Index"
                                        title="Customer Satisfaction Index"
                                        chartType="ComboChart"
                                        data={CSITrendData}
                                        options={options}
                                        width="100%" height="240px" />
                                </div>
                            </div>
                        )}



                        {(dialogPivotKey === 'Effort Variation' || dialogPivotKey === 'Effort Variation_Month') && (
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Effort Variation trend</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-ev`} chartKey="Effort Variation" title="Effort Variation" chartType="ComboChart" data={efforVatiationChartData ?? []} options={options} width="100%" height="240px" />
                                </div>
                            </div>
                        )}

                        {(dialogPivotKey === 'Schedule Variation' || dialogPivotKey === 'Schedule Variation_Month') && (
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Schedule Variation trend</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-sv`} chartKey="Schedule Variation" title="Schedule Variation" chartType="ComboChart" data={scheduledVatiationChartData ?? []} options={options} width="100%" height="240px" />
                                </div>
                            </div>
                        )}

                        {(dialogPivotKey === 'Overall Productivity' || dialogPivotKey === 'Overall Productivity_Month') && (<>
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Overall Productivity trend</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-oap`} chartKey="Overall Productivity" title="Overall Productivity" chartType="ComboChart" data={OverallProductivityChartData ?? []} options={options} width="100%" height="240px" />
                                </div>
                            </div>
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Effort Distribution</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-effortdist`} chartKey="EffortDistribution" title="Effort Distribution" chartType="PieChart" data={EffortDistributionData ?? []} options={{ legend: { position: 'right' }, pieSliceText: 'value', pieSliceTextStyle: { fontSize: 12, color: '#fff' }, tooltip: { trigger: 'focus', text: 'value' }, slices: { 0: { color: '#e91e63' }, 1: { color: '#ff7043' }, 2: { color: '#ffc107' }, 3: { color: '#8bc34a' } } }} width="100%" height="220px" />
                                </div>
                            </div>
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Requirement Analysis Effort Density</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-raed`} chartKey="RequirementAnalysisEffortDensity" title="Requirement Analysis Effort Density" chartType="ComboChart" data={RequirementAnalysisEffortDensityChartData ?? []} options={options} width="100%" height="240px" />
                                </div>
                            </div>
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Coding Productivity</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-ded`} chartKey="CodingProductivity" title="Coding Productivity" chartType="ComboChart" data={CodingProductivityChartData ?? []} options={options} width="100%" height="240px" />
                                </div>
                            </div>
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Coding Review Effort Density</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-ded`} chartKey="CodingReviewEffortDensity" title="Coding Review Effort Density" chartType="ComboChart" data={CodeReviewEffortDensityChartData ?? []} options={options} width="100%" height="240px" />
                                </div>
                            </div>

                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Code Rework Effort Density</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-cred`} chartKey="CodeReworkEffortDensity" title="Code Rework Effort Density" chartType="ComboChart" data={CodeReworkEffortDensityChartData ?? []} options={options} width="100%" height="240px" />
                                </div>
                            </div>

                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Unit Testing Effort Density</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-ted`} chartKey="UnitTestingEffortDensity" title="Unit Testing Effort Density" chartType="ComboChart" data={UnitTestingEffortDensityChartData ?? []} options={options} width="100%" height="240px" />
                                </div>
                            </div>

                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Test Execution Effort Density</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-ted`} chartKey="TestExecutionEffortDensity" title="Test Execution Effort Density" chartType="ComboChart" data={TestingExecutionEffortDensityChartData ?? []} options={options} width="100%" height="240px" />
                                </div>
                            </div>
                        </>
                        )}
                        {
                            dialogPivotKey === 'Resource Utilization' && (
                                <div style={chartsContainerStyle}>
                                    <div style={chartBoxStyle}>
                                        <h3>Resource Utilization trend</h3>
                                        <ClickableChart key={`dlg-${dialogChartKey}-ru`} chartKey="Resource Utilization" title="Resource Utilization (%)" chartType="ComboChart" data={ResourceUtilizationTrendChartData ?? []} options={options} width="100%" height="240px" />
                                    </div>
                                </div>
                            )
                        }
                        {
                            dialogPivotKey === 'Cost of Quality' && (
                                <div style={chartsContainerStyle}>
                                    <div style={chartBoxStyle}>
                                        <h3>Cost of Quality</h3>
                                        <ClickableChart key={`dlg-${dialogChartKey}-ru`} chartKey="Cost of Quality" title="Cost of Quality (%)" chartType="ComboChart" data={CostOfQualityTrendChartData ?? []} options={options} width="100%" height="240px" />
                                    </div>
                                </div>
                            )
                        }
                        {
                            (dialogPivotKey === 'Defect density' || dialogPivotKey === 'Internal Defects') && (
                                <>
                                    <div style={{ ...chartsContainerStyle, minHeight: 220 }}>
                                        <div style={{ ...chartBoxStyle, paddingBottom: 12 }}>
                                            <h3>Internal Defects</h3>
                                            <DetailsList
                                                items={defectsDetailsListItems}
                                                columns={defectsDetailsListColumns}
                                                selectionMode={SelectionMode.none}
                                                checkboxVisibility={CheckboxVisibility.hidden}
                                            />
                                        </div>
                                    </div>
                                    <div style={chartsContainerStyle}>
                                        <div style={chartBoxStyle}>
                                            <h3>Code Review Defect Density</h3>
                                            <ClickableChart key={`dlg-${dialogChartKey}-crdd`} chartKey="CodeReviewDefectDensity" title="Code Review Defect Density" chartType="ComboChart" data={CodeReviewDefectDensityChartData ?? []} options={options} width="100%" height="240px" />
                                        </div>
                                    </div>
                                    <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Defect Density</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-pddd`} chartKey="DefectDensity" title="Defect Density" chartType="ComboChart" data={TestingPhaseDefectDensityChartData ?? []} options={options} width="100%" height="240px" />
                                </div>
                            </div>
                                </>
                            )
                        }

                        {(dialogPivotKey === 'Delivered Defect Density' || dialogPivotKey === 'Post Delivery Defects') && (
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <h3>Post Delivery Defect Density</h3>
                                    <ClickableChart key={`dlg-${dialogChartKey}-pddd`} chartKey="PostDeliveryDefectDensity" title="Post Delivery Defect Density" chartType="ComboChart" data={PostDeliveryDefectDensityChartData ?? []} options={options} width="100%" height="240px" />
                                </div>
                            </div>
                        )}

                    </div>
                    <DialogFooter>
                        <PrimaryButton onClick={closeDialog} text="Close" />
                    </DialogFooter>
                </Dialog>

                {/* <Pivot
                    aria-label="Metrics Tabs"
                    selectedKey={selectedPivotKey}
                    onLinkClick={(item) => setSelectedPivotKey(item?.props.itemKey ?? 'Customer')}
                    styles={{ root: { width: '100%' } }}
                >
                    <PivotItem itemKey="Customer" headerText="Customer Satisfaction Index">
                        {/* Two charts side-by-side: rating with limits + risk 
                        <div style={chartsContainerStyle}>
                            <div style={chartBoxStyle}>
                                <ClickableChart chartKey="rating"
                                    title="Customer Rating"
                                    chartType="ComboChart"
                                    data={CSITrendData}
                                    options={options}
                                    width="100%" height="240px" />
                            </div>
                            {/* <div style={chartBoxStyle}>
                                    <ClickableChart chartKey="risk" title="Risks Open" chartType="AreaChart" data={riskData} options={{ colors: ['#f39c12'] }} width="100%" height="240px" />
                                </div> 
                        </div>
                    </PivotItem>

                    <PivotItem itemKey="EffortDistribution" headerText="Effort Distribution">
                        {/* Two charts side-by-side: rating with limits + risk 
                        <div style={chartsContainerStyle}>
                            <div style={chartBoxStyle}>
                                <ClickableChart chartKey="EffortDistribution"
                                    title="Effort Distribution" chartType="PieChart"
                                    data={EffortDistributionData ?? []}
                                    options={{

                                        legend: { position: 'right' },

                                        pieSliceText: 'value', // alternatives: 'percentage' | 'label' | 'none'
                                        pieSliceTextStyle: { fontSize: 12, color: '#fff' },

                                        // Optional: format tooltip to show values
                                        tooltip: { trigger: 'focus', text: 'value' },

                                        slices: {
                                            0: { color: '#e91e63' },
                                            1: { color: '#ff7043' },
                                            2: { color: '#ffc107' },
                                            3: { color: '#8bc34a' }
                                        }
                                    }}
                                    width="100%"
                                    height="220px" />

                            </div>
                            {/* <div style={chartBoxStyle}>
                                    <ClickableChart chartKey="risk" title="Risks Open" chartType="AreaChart" data={riskData} options={{ colors: ['#f39c12'] }} width="100%" height="240px" />
                                </div> *
                        </div>
                    </PivotItem>

                    <PivotItem itemKey="Effort Variation" headerText="Effort Variation">
                        {/* Two charts side-by-side: rating with limits + risk 
                        <div style={chartsContainerStyle}>
                            <div style={chartBoxStyle}>
                                <ClickableChart chartKey="Effort Variation"
                                    title="Effort Variation"
                                    chartType="ComboChart"
                                    data={efforVatiationChartData ?? []}
                                    options={options}
                                    width="100%" height="240px" />
                            </div>
                            {/* <div style={chartBoxStyle}>
                                    <ClickableChart chartKey="risk" title="Risks Open" chartType="AreaChart" data={riskData} options={{ colors: ['#f39c12'] }} width="100%" height="240px" />
                                </div> 
                        </div>
                    </PivotItem>

                    <PivotItem itemKey="Schedule Variation" headerText="Schedule Variation">
                        {/* Two charts side-by-side: rating with limits + risk 
                        <div style={chartsContainerStyle}>
                            <div style={chartBoxStyle}>
                                <ClickableChart chartKey="Schedule Variation"
                                    title="Schedule Variation"
                                    chartType="ComboChart"
                                    data={scheduledVatiationChartData ?? []}
                                    options={options}
                                    width="100%" height="240px" />
                            </div>
                            {/* <div style={chartBoxStyle}>
                                    <ClickableChart chartKey="risk" title="Risks Open" chartType="AreaChart" data={riskData} options={{ colors: ['#f39c12'] }} width="100%" height="240px" />
                                </div> 
                        </div>
                    </PivotItem>

                    <PivotItem itemKey="Overall Productivity" headerText="Overall Productivity">
                        {/* Two charts side-by-side: rating with limits + risk 
                        <div style={chartsContainerStyle}>
                            <div style={chartBoxStyle}>
                                <ClickableChart chartKey="Overall Productivity"
                                    title="Overall Productivity"
                                    chartType="ComboChart"
                                    data={OverallProductivityChartData ?? []}
                                    options={options}
                                    width="100%" height="240px" />
                            </div>
                            {/* <div style={chartBoxStyle}>
                                    <ClickableChart chartKey="risk" title="Risks Open" chartType="AreaChart" data={riskData} options={{ colors: ['#f39c12'] }} width="100%" height="240px" />
                                </div> 
                        </div>
                    </PivotItem>

                    {/* <PivotItem headerText="Velocity" itemKey='Velocity'>
                            {/* Two charts side-by-side: velocity with productivity limits + defect with limits 
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <ClickableChart
                                        chartKey="velocity"
                                        title="Velocity"
                                        chartType="ComboChart"
                                        data={velocityData}
                                        options={{
                                            legend: { position: 'none' },
                                            seriesType: 'line',
                                            series: {
                                                1: { color: '#28a745', lineWidth: 1, lineDashStyle: [4, 4] },
                                                2: { color: '#dc3545', lineWidth: 1, lineDashStyle: [4, 4] },
                                            },
                                            colors: ['#007bff'],
                                        }}
                                        width="100%"
                                        height="240px"
                                    />
                                </div>
                                <div style={chartBoxStyle}>
                                    <ClickableChart
                                        chartKey="defect"
                                        title="Defect Density"
                                        chartType="ComboChart"
                                        data={defectData}
                                        options={{
                                            seriesType: 'bars',
                                            series: {
                                                1: { type: 'line', color: '#28a745', lineWidth: 1, lineDashStyle: [4, 4] },
                                                2: { type: 'line', color: '#dc3545', lineWidth: 1, lineDashStyle: [4, 4] },
                                            },
                                            colors: ['#dc3545'],
                                            legend: { position: 'none' },
                                        }}
                                        width="100%"
                                        height="240px"
                                    />
                                </div>
                            </div>
                        </PivotItem>

                        <PivotItem itemKey="Defect" headerText="Defect Density">
                            {/* Two charts side-by-side 
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <ClickableChart chartKey="defect_table" title="Defect Density (bar)" chartType="ColumnChart" data={defectData} options={{ colors: ['#dc3545'] }} width="100%" height="240px" />
                                </div>
                                <div style={chartBoxStyle}>
                                    <ClickableChart
                                        chartKey="week"
                                        title="Weekly Report"
                                        chartType="LineChart"
                                        data={weekData}
                                        options={{ colors: ['#28a745'] }}
                                        width="100%"
                                        height="240px"
                                    />
                                </div>
                            </div>
                        </PivotItem>
                        

                        <PivotItem itemKey="Resource" headerText="Resource Utilization">
                            {/* Two charts side-by-side 
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <ClickableChart chartKey="rating_2" title="Resource Utilization" chartType="LineChart" data={resourceData} options={{ colors: ['#28a745'] }} width="100%" height="240px" />
                                </div>
                                <div style={chartBoxStyle}>
                                    <ClickableChart chartKey="risk_2" title="Risks Open" chartType="AreaChart" data={riskData} options={{ colors: ['#f39c12'] }} width="100%" height="240px" />
                                </div>
                            </div>
                        </PivotItem>

                        <PivotItem itemKey="Weekly" headerText="Weekly Report">
                            {/* Two charts side-by-side (weekly with limits + complementary) 
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <ClickableChart chartKey="week_2" title="Weekly Report" chartType="ComboChart" data={weekData} options={{ legend: { position: 'none' }, seriesType: 'line', series: { 1: { color: '#28a745', lineWidth: 1, lineDashStyle: [4, 4] }, 2: { color: '#dc3545', lineWidth: 1, lineDashStyle: [4, 4] } }, colors: ['#6f42c1'] }} width="100%" height="300px" />
                                </div>
                                <div style={chartBoxStyle}>
                                    <ClickableChart chartKey="velocity_2" title="Velocity (bar)" chartType="ColumnChart" data={velocityData} options={{ colors: ['#007bff'] }} width="100%" height="300px" />
                                </div>
                            </div>
                        </PivotItem>

                        <PivotItem itemKey="Risk" headerText="Risk">
                            {/* Two charts side-by-side 
                            <div style={chartsContainerStyle}>
                                <div style={chartBoxStyle}>
                                    <ClickableChart chartKey="risk_3" title="Risks Open (line)" chartType="LineChart" data={riskData} options={{ colors: ['#e91e63'] }} width="100%" height="300px" />
                                </div>
                                <div style={chartBoxStyle}>
                                    <ClickableChart chartKey="defect_2" title="Defect Area" chartType="AreaChart" data={defectData} options={{ colors: ['#f39c12'] }} width="100%" height="300px" />
                                    <div style={{ marginTop: 12 }}>
                                        <DefaultButton text="Open Risk Report" />
                                    </div>
                                </div>
                            </div>
                        </PivotItem> 
                </Pivot> */}



                {/* Critical Risk and Findings Summary side-by-side */}
                <section style={belowNavContainerStyle}>
                    <div style={halfChartBoxStyle}>
                        <Text variant="large" styles={{ root: { marginBottom: 8 } }}>Risk Summary</Text>
                        <ClickableChart chartKey="Risk Summary" title="Risk Summary" chartType="ColumnChart" data={riskChartData} options={{
                            legend: { position: 'none' },
                            colors: ['#42a5f5'],
                            hAxis: { title: 'Risk Exposure' },
                            vAxis: { title: 'Number of Risks', minValue: 0, format: '0', gridlines: { color: '#eee' } },
                            bar: { groupWidth: '60%' },
                            annotations: {
                                alwaysOutside: true,
                                textStyle: { fontSize: 12, color: '#333' }
                            }
                        }} width="100%" height="260px" />
                        {/* 
                        <div style={{ marginTop: 10 }}>
                            <Text variant="small">Total critical findings open: <strong>158</strong></Text>
                            <div style={{ marginTop: 8 }}>
                              
                                <ClickableChart chartKey="criticalBreakdown" title="Critical Breakdown" chartType="PieChart" data={criticalBreakdownData} options={{ legend: { position: 'right' }, colors: ['#9c27b0', '#03a9f4'] }} width="100%" height="140px" />
                            </div>
                        </div> */}
                    </div>

                    <div style={halfChartBoxStyle}>
                        <Text variant="large" styles={{ root: { marginBottom: 8 } }}>Findings Summary</Text>
                        <div>
                            {/* <Chart
                                chartType="PieChart"
                                width="100%"
                                height="180px"
                                data={findingsTypeData}
                                options={{
                                    pieHole: 0.5,
                                    legend: { position: 'right' },
                                    slices: { 0: { color: '#d32f2f' }, 1: { color: '#1976d2' }, 2: { color: '#388e3c' } },
                                }}
                            /> */}
                            <ClickableChart chartKey="findings" title="Findings by Type" chartType="PieChart" data={facilitationReportDataForChart} options={{ pieHole: 0, legend: { position: 'right' }, pieSliceTextStyle: { fontSize: 12, color: '#fff' }, slices: { 0: { color: '#d32f2f' }, 1: { color: '#1976d2' }, 2: { color: '#388e3c' } } }} width="100%" height="180px" />
                        </div>

                        {/* <div style={{ marginTop: 10 }}>
                            <Text variant="small">Top contributors (departments & projects)</Text>
                            <div style={{ marginTop: 8 }}>
                                {/* <Chart
                                    chartType="ColumnChart"
                                    width="100%"
                                    height="180px"
                                    data={contributorsData}
                                    options={{
                                        legend: { position: 'none' },
                                        colors: ['#5c6bc0'],
                                        hAxis: { textStyle: { fontSize: 10 } },
                                        vAxis: { minValue: 0 },
                                    }}
                                /> 
                                <ClickableChart chartKey="contributors" title="Top Contributors" chartType="ColumnChart" data={contributorsData} options={{ legend: { position: 'none' }, colors: ['#5c6bc0'], hAxis: { textStyle: { fontSize: 10 } }, vAxis: { minValue: 0 } }} width="100%" height="180px" />
                            </div>
                        </div> */}

                        {/* <div style={{ marginTop: 8 }}>
                            <Text variant="small">Details:</Text>
                            <ul style={{ marginTop: 6 }}>
                                <li>Total open findings: <strong>736</strong></li>
                                <li>NCs: <strong>101</strong>, Observations: <strong>356</strong>, Facilitation: <strong>279</strong></li>
                                <li>Top departments: <strong>IT IS (43)</strong>, <strong>IT IS Europe (35)</strong>, <strong>FPA (27)</strong></li>
                                <li>Top projects: <strong>Dormer (39)</strong>, <strong>Americhem AI platform (23)</strong>, <strong>RLUS Analytics (23)</strong></li>
                            </ul>
                        </div> */}
                    </div>
                </section>

                {/* NEW: Aging findings chart (full-width within content wrapper) */}
                <section style={belowNavContainerStyle}>
                    <div style={halfChartBoxStyle}>
                        <Text variant="large" styles={{ root: { marginBottom: 8 } }}>Aging Findings</Text>
                        <ClickableChart chartKey="aging" title="Aging Findings" chartType="ColumnChart" data={agingFindingsData} options={{ legend: { position: 'none' }, colors: ['#42a5f5'], hAxis: { title: 'Age bucket' }, vAxis: { title: 'Number of findings', minValue: 0, format: '0', gridlines: { color: '#eee' } }, bar: { groupWidth: '60%' } }} width="100%" height="260px" />
                        {/* <div style={{ marginTop: 10 }}>
                            <Text variant="small">Summary: <strong>736</strong> open findings across aging buckets.</Text>
                        </div> */}
                    </div>
                    <div style={halfChartBoxStyle}>
                        <Text variant="large" styles={{ root: { marginBottom: 8 } }}>PCI</Text>
                        <ClickableChart
                            chartKey="PCI"
                            title="Process Compliance Index (PCI)"
                            chartType="LineChart"
                            data={pciChartData}
                            options={pciOptions}
                            width="100%"
                            height="260px"
                        />
                        {/* <div style={{ marginTop: 10 }}>
                            <Text variant="small">Summary: <strong>736</strong> open findings across aging buckets.</Text>
                        </div> */}
                    </div>
                </section>
                <section style={{ margin: '20px 0' }}>
                    <div style={kpiAreaStyle}>
                        {/* Open items KPI cards */}
                        <div style={{ ...kpiCardStyle(), cursor: 'default', border: '1px solid rgba(16,24,40,0.04)' }}>
                            <div>
                                <Text variant="small" styles={{ root: { color: '#666' } }}>Open Root Cause</Text>
                                <div style={{ fontSize: 26, fontWeight: 700, marginTop: 8, marginBottom: 6, color: '#d9534f' }}>{RCAItems.filter(i => i?.ActualClosureDateCorrection == "" || i?.ActualClosureDateCorrective == "" || i?.ActualClosureDatePreventive == "").length}</div>
                            </div>
                        </div>
                        <div style={{ ...kpiCardStyle(), cursor: 'default', border: '1px solid rgba(16,24,40,0.04)' }}>
                            <div>
                                <Text variant="small" styles={{ root: { color: '#666' } }}>Open Issues</Text>
                                <div style={{ fontSize: 26, fontWeight: 700, marginTop: 8, marginBottom: 6, color: '#f0ad4e' }}>{RAIDdata.filter(i => i?.SelectType === 'Issue' && i?.RiskStatus == "In Progress").length}</div>
                            </div>
                        </div>
                        <div style={{ ...kpiCardStyle(), cursor: 'default', border: '1px solid rgba(16,24,40,0.04)' }}>
                            <div>
                                <Text variant="small" styles={{ root: { color: '#666' } }}>Open Action Items</Text>
                                <div style={{ fontSize: 26, fontWeight: 700, marginTop: 8, marginBottom: 6, color: '#5bc0de' }}>{0}</div>
                            </div>
                        </div>
                    </div>
                </section>
            </div>
        </div>
    );
}



