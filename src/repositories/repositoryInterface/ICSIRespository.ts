//IProjectMetricsRepository
import { WebPartContext } from '@microsoft/sp-webpart-base';
import IGenericService from '../../services/IGenericServices';
import {  ICSI } from '../../Models/ICSI';

export interface ICSIRepository {
  
  refresh(): void;
  getCacheStatus(): { cached: boolean; itemCount: number; age: number };
  setService(service: IGenericService): void;
  getCSIValues(useCache?: boolean, context?: WebPartContext, selectedStartDate?: string, selectedEndDate?: string): Promise<ICSI[]>;}

export default ICSIRepository;