import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPPOApprovers } from '../../Models/IPPOApprovers';
import IGenericService from '../../services/IGenericServices';

export interface IPPOApproversRepository {
	/**
	 * Fetch all PPOApprovers items. Returns array of IPPOApprovers.
	 */
	getAllItems(useCache?: boolean, context?: WebPartContext): Promise<IPPOApprovers[]>;

	/**
	 * Clear internal cache
	 */
	refresh(): void;

	getCacheStatus(): { cached: boolean; itemCount: number; age: number };

	setService(service: IGenericService): void;
}

export default IPPOApproversRepository;