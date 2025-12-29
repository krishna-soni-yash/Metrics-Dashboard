export interface IPPOApprovers {
    ID: number;
    LinkTitle: string;
    SiteURL: string;
    Reviewer: string;
    ReviewerEmail?: string;
    BUH: string;
    BUHEmail?: string;
    ProjectManager: string;
    ProjectManagerEmail?: string;
    isCurrentReviewer?: boolean;
    isCurrentBUH?: boolean;
    isCurrentProjectManager?: boolean;
    currentUserRoles?: string[];
    teamSize?: number;
}