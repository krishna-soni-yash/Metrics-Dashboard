//IProjectMetricsRepository
import { WebPartContext } from '@microsoft/sp-webpart-base';
import IGenericService from '../../services/IGenericServices';
import {  IFacilitationReport } from '../../Models/IFacilationReport';

export interface IFacilitationReportRepository {
  
  refresh(): void;
  getCacheStatus(): { cached: boolean; itemCount: number; age: number };
  setService(service: IGenericService): void;
  getFacilitationValues(useCache?: boolean, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<IFacilitationReport[]>;}

export default IFacilitationReportRepository;