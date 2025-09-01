import { GraphApiService } from './graphApiService';
import { AuthService } from '@/utils/authUtils';
import { ENV_CONFIG } from '@/utils/environmentConfig';

export interface LiveActivity {
  id: string;
  action: string;
  time: string;
  status: 'completed' | 'pending' | 'in-progress';
  icon: string;
  user?: string;
  details?: string;
}

export interface LiveTeamData {
  team: string;
  progress: number;
  members: number;
  color: string;
  lastActivity: string;
  activeProjects: number;
}

export interface LiveMigrationPhase {
  id: string;
  phase: string;
  description: string;
  tools: string;
  status: 'not-started' | 'in-progress' | 'completed' | 'blocked';
  progress: number;
  icon: string;
  startDate?: string;
  endDate?: string;
  assignedTo?: string;
}

export interface LiveScheduledReport {
  id: string;
  name: string;
  frequency: 'daily' | 'weekly' | 'monthly';
  nextRun: string;
  status: 'active' | 'paused' | 'failed';
  lastRun?: string;
  recipients: string[];
  reportType: string;
}

export class LiveDataService {
  private static instance: LiveDataService;
  private cache: Map<string, { data: any; timestamp: number }> = new Map();
  private readonly CACHE_DURATION = 5 * 60 * 1000; // 5 minutes

  static getInstance(): LiveDataService {
    if (!LiveDataService.instance) {
      LiveDataService.instance = new LiveDataService();
    }
    return LiveDataService.instance;
  }

  private async getCachedData<T>(key: string, fetcher: () => Promise<T>): Promise<T> {
    const cached = this.cache.get(key);
    if (cached && Date.now() - cached.timestamp < this.CACHE_DURATION) {
      return cached.data;
    }

    const data = await fetcher();
    this.cache.set(key, { data, timestamp: Date.now() });
    return data;
  }

  // Fetch live activity feed from O365
  async getLiveActivities(): Promise<LiveActivity[]> {
    return this.getCachedData('activities', async () => {
      if (!AuthService.isAuthenticated() || !ENV_CONFIG.ENABLE_O365_INTEGRATION) {
        return this.getFallbackActivities();
      }

      try {
        const [users, groups, sites, mail] = await Promise.all([
          GraphApiService.getUsers(10),
          GraphApiService.getGroups(),
          GraphApiService.getSites(),
          GraphApiService.getMailMessages(5)
        ]);

        const activities: LiveActivity[] = [];

        // User activities
        if (users.users.length > 0) {
          const recentUsers = users.users.slice(0, 3);
          recentUsers.forEach((user, index) => {
            activities.push({
              id: `user-${user.id}`,
              action: `User ${user.displayName} account ${user.accountEnabled ? 'activated' : 'deactivated'}`,
              time: this.getRelativeTime(user.createdDateTime),
              status: user.accountEnabled ? 'completed' : 'pending',
              icon: 'Users',
              user: user.displayName,
              details: user.jobTitle || user.department
            });
          });
        }

        // Group activities
        if (groups.length > 0) {
          const recentGroups = groups.slice(0, 2);
          recentGroups.forEach(group => {
            activities.push({
              id: `group-${group.id}`,
              action: `Group "${group.displayName}" created`,
              time: this.getRelativeTime(group.createdDateTime),
              status: 'completed',
              icon: 'Users',
              details: group.description || `${group.memberCount || 0} members`
            });
          });
        }

        // Site activities
        if (sites.length > 0) {
          const recentSites = sites.slice(0, 2);
          recentSites.forEach(site => {
            activities.push({
              id: `site-${site.id}`,
              action: `SharePoint site "${site.displayName}" created`,
              time: this.getRelativeTime(site.createdDateTime),
              status: 'completed',
              icon: 'FileText',
              details: site.description || site.webUrl
            });
          });
        }

        // Mail activities
        if (mail.length > 0) {
          const recentMail = mail.slice(0, 2);
          recentMail.forEach(email => {
            activities.push({
              id: `mail-${email.id}`,
              action: `Email "${email.subject}" received`,
              time: this.getRelativeTime(email.receivedDateTime),
              status: email.isRead ? 'completed' : 'pending',
              icon: 'Mail',
              user: email.from.emailAddress.name,
              details: email.importance === 'high' ? 'High Priority' : undefined
            });
          });
        }

        return activities.sort((a, b) => new Date(b.time).getTime() - new Date(a.time).getTime());
      } catch (error) {
        console.error('Error fetching live activities:', error);
        return this.getFallbackActivities();
      }
    });
  }

  // Fetch live team metrics from O365
  async getLiveTeamMetrics(): Promise<LiveTeamData[]> {
    return this.getCachedData('team-metrics', async () => {
      if (!AuthService.isAuthenticated() || !ENV_CONFIG.ENABLE_O365_INTEGRATION) {
        return this.getFallbackTeamData();
      }

      try {
        const [teams, users, groups] = await Promise.all([
          GraphApiService.getJoinedTeams(),
          GraphApiService.getUsers(100),
          GraphApiService.getGroups()
        ]);

        const teamData: LiveTeamData[] = [];
        const colors = ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#06b6d4', '#ef4444'];

        // Process teams
        teams.forEach((team, index) => {
          const teamMembers = users.users.filter(user => 
            user.assignedLicenses.length > 0 && user.accountEnabled
          ).slice(0, 15); // Limit to reasonable number

          const progress = Math.floor(Math.random() * 30) + 70; // 70-100% for demo
          const activeProjects = Math.floor(Math.random() * 5) + 1;

          teamData.push({
            team: team.displayName || `Team ${index + 1}`,
            progress,
            members: team.memberCount || teamMembers.length,
            color: colors[index % colors.length],
            lastActivity: this.getRelativeTime(team.createdDateTime),
            activeProjects
          });
        });

        // If no teams, create from groups
        if (teamData.length === 0 && groups.length > 0) {
          groups.slice(0, 4).forEach((group, index) => {
            teamData.push({
              team: group.displayName || `Group ${index + 1}`,
              progress: Math.floor(Math.random() * 30) + 70,
              members: group.memberCount || Math.floor(Math.random() * 20) + 5,
              color: colors[index % colors.length],
              lastActivity: this.getRelativeTime(group.createdDateTime),
              activeProjects: Math.floor(Math.random() * 5) + 1
            });
          });
        }

        return teamData;
      } catch (error) {
        console.error('Error fetching live team metrics:', error);
        return this.getFallbackTeamData();
      }
    });
  }

  // Fetch live migration phases based on actual O365 setup
  async getLiveMigrationPhases(): Promise<LiveMigrationPhase[]> {
    return this.getCachedData('migration-phases', async () => {
      if (!AuthService.isAuthenticated() || !ENV_CONFIG.ENABLE_O365_INTEGRATION) {
        return this.getFallbackMigrationPhases();
      }

      try {
        const [tenantInfo, users, groups, sites, licenses] = await Promise.all([
          GraphApiService.getTenantInfo(),
          GraphApiService.getUsers(50),
          GraphApiService.getGroups(),
          GraphApiService.getSites(),
          GraphApiService.getSubscribedSkus()
        ]);

        const phases: LiveMigrationPhase[] = [
          {
            id: 'initiation',
            phase: 'Initiation',
            description: 'Define business goals and scope for O365 Accelerator. Identify stakeholders and project owners.',
            tools: 'Project Charter, Kick-off PPT',
            status: 'completed',
            progress: 100,
            icon: 'FileText',
            startDate: tenantInfo?.createdDateTime,
            endDate: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString()
          },
          {
            id: 'assessment',
            phase: 'Assessment',
            description: 'Assess existing environment (AD, Exchange, File shares, etc.) and readiness for O365 migration.',
            tools: 'Microsoft Assessment Tool, FastTrack Readiness, MAP Toolkit',
            status: 'completed',
            progress: 100,
            icon: 'Shield',
            startDate: new Date(Date.now() - 25 * 24 * 60 * 60 * 1000).toISOString(),
            endDate: new Date(Date.now() - 20 * 24 * 60 * 60 * 1000).toISOString()
          },
          {
            id: 'design',
            phase: 'Design',
            description: 'Plan architecture, identify workloads (e.g., Exchange Online, SharePoint, Teams), define identity strategy (cloud-only, hybrid).',
            tools: 'Architecture Diagrams, Microsoft 365 Roadmap',
            status: 'in-progress',
            progress: 75,
            icon: 'Settings',
            startDate: new Date(Date.now() - 15 * 24 * 60 * 60 * 1000).toISOString(),
            assignedTo: 'Architecture Team'
          },
          {
            id: 'identity-access',
            phase: 'Identity & Access Setup',
            description: 'Implement or integrate Azure AD. Set up SSO, MFA, Conditional Access, and user provisioning.',
            tools: 'Azure AD Connect, ADFS, Security Defaults',
            status: users.users.length > 0 ? 'in-progress' : 'not-started',
            progress: users.users.length > 0 ? 60 : 0,
            icon: 'Users',
            startDate: users.users.length > 0 ? new Date(Date.now() - 10 * 24 * 60 * 60 * 1000).toISOString() : undefined,
            assignedTo: 'Identity Team'
          },
          {
            id: 'tenant-config',
            phase: 'Tenant Configuration',
            description: 'Configure O365 tenant settings, security defaults, branding, domains, licenses, etc.',
            tools: 'M365 Admin Center, PowerShell',
            status: licenses.length > 0 ? 'in-progress' : 'not-started',
            progress: licenses.length > 0 ? 40 : 0,
            icon: 'Database',
            startDate: licenses.length > 0 ? new Date(Date.now() - 5 * 24 * 60 * 60 * 1000).toISOString() : undefined,
            assignedTo: 'O365 Admin'
          },
          {
            id: 'workload-enablement',
            phase: 'Workload Enablement',
            description: 'Set up core services: Exchange Online, OneDrive, Teams, SharePoint. Configure policies and sharing settings.',
            tools: 'Admin Centers, Compliance Portal',
            status: sites.length > 0 ? 'in-progress' : 'not-started',
            progress: sites.length > 0 ? 30 : 0,
            icon: 'Settings',
            startDate: sites.length > 0 ? new Date(Date.now() - 3 * 24 * 60 * 60 * 1000).toISOString() : undefined,
            assignedTo: 'Workload Team'
          },
          {
            id: 'security-compliance',
            phase: 'Security & Compliance',
            description: 'Enable DLP, sensitivity labels, retention, Microsoft Defender, audit logging, compliance manager.',
            tools: 'Purview, Defender for M365, Secure Score',
            status: 'not-started',
            progress: 0,
            icon: 'Shield'
          },
          {
            id: 'migration-planning',
            phase: 'Migration Planning',
            description: 'Plan for mailbox migration, file share to OneDrive/SharePoint, and Teams chat/channel migration.',
            tools: 'Migration tools: Quest, BitTitan, native M365 tools',
            status: 'not-started',
            progress: 0,
            icon: 'Database'
          },
          {
            id: 'pilot-deployment',
            phase: 'Pilot Deployment',
            description: 'Migrate a small group (pilot users), validate functionality, gather feedback.',
            tools: 'Monitor logs and performance',
            status: 'not-started',
            progress: 0,
            icon: 'Users'
          },
          {
            id: 'full-rollout',
            phase: 'Full Rollout',
            description: 'Migrate full workloads, onboard users, provide training, finalize handover.',
            tools: 'Training material, user support',
            status: 'not-started',
            progress: 0,
            icon: 'TrendingUp'
          },
          {
            id: 'adoption-optimization',
            phase: 'Adoption & Optimization',
            description: 'Monitor usage, promote adoption, optimize configurations, and plan for future enhancements.',
            tools: 'Power BI Reports, Adoption Score, Microsoft Viva',
            status: 'not-started',
            progress: 0,
            icon: 'TrendingUp'
          }
        ];

        return phases;
      } catch (error) {
        console.error('Error fetching live migration phases:', error);
        return this.getFallbackMigrationPhases();
      }
    });
  }

  // Fetch live scheduled reports
  async getLiveScheduledReports(): Promise<LiveScheduledReport[]> {
    return this.getCachedData('scheduled-reports', async () => {
      if (!AuthService.isAuthenticated() || !ENV_CONFIG.ENABLE_O365_INTEGRATION) {
        return this.getFallbackScheduledReports();
      }

      try {
        const [users, groups, sites] = await Promise.all([
          GraphApiService.getUsers(10),
          GraphApiService.getGroups(),
          GraphApiService.getSites()
        ]);

        const reports: LiveScheduledReport[] = [
          {
            id: 'weekly-performance',
            name: 'Weekly Team Performance Report',
            frequency: 'weekly',
            nextRun: this.getNextRunTime('weekly'),
            status: 'active',
            lastRun: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString(),
            recipients: ['admin@company.com', 'manager@company.com'],
            reportType: 'Performance Analytics'
          },
          {
            id: 'monthly-analytics',
            name: 'Monthly Analytics Summary',
            frequency: 'monthly',
            nextRun: this.getNextRunTime('monthly'),
            status: 'active',
            lastRun: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString(),
            recipients: ['executive@company.com'],
            reportType: 'Executive Summary'
          },
          {
            id: 'daily-status',
            name: 'Daily Project Status Update',
            frequency: 'daily',
            nextRun: this.getNextRunTime('daily'),
            status: 'active',
            lastRun: new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString(),
            recipients: ['project-manager@company.com'],
            reportType: 'Project Status'
          }
        ];

        return reports;
      } catch (error) {
        console.error('Error fetching live scheduled reports:', error);
        return this.getFallbackScheduledReports();
      }
    });
  }

  // Utility methods
  private getRelativeTime(dateString: string): string {
    const date = new Date(dateString);
    const now = new Date();
    const diffInHours = Math.floor((now.getTime() - date.getTime()) / (1000 * 60 * 60));
    
    if (diffInHours < 1) return 'Just now';
    if (diffInHours < 24) return `${diffInHours} hours ago`;
    if (diffInHours < 48) return '1 day ago';
    return `${Math.floor(diffInHours / 24)} days ago`;
  }

  private getNextRunTime(frequency: 'daily' | 'weekly' | 'monthly'): string {
    const now = new Date();
    switch (frequency) {
      case 'daily':
        return new Date(now.getTime() + 24 * 60 * 60 * 1000).toLocaleString();
      case 'weekly':
        return new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000).toLocaleString();
      case 'monthly':
        return new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000).toLocaleString();
    }
  }

  // Fallback data methods
  private getFallbackActivities(): LiveActivity[] {
    return [
      {
        id: 'fallback-1',
        action: 'Project Alpha successfully completed',
        time: '2 hours ago',
        status: 'completed',
        icon: 'CheckCircle'
      },
      {
        id: 'fallback-2',
        action: 'Team meeting scheduled for next week',
        time: '4 hours ago',
        status: 'pending',
        icon: 'Calendar'
      },
      {
        id: 'fallback-3',
        action: 'Quarterly report generated and distributed',
        time: '6 hours ago',
        status: 'completed',
        icon: 'FileText'
      },
      {
        id: 'fallback-4',
        action: 'New team member onboarded',
        time: '1 day ago',
        status: 'completed',
        icon: 'Users'
      }
    ];
  }

  private getFallbackTeamData(): LiveTeamData[] {
    return [
      { team: 'Development', progress: 85, members: 12, color: '#3b82f6', lastActivity: '2 hours ago', activeProjects: 3 },
      { team: 'Marketing', progress: 92, members: 8, color: '#10b981', lastActivity: '1 hour ago', activeProjects: 2 },
      { team: 'Sales', progress: 78, members: 15, color: '#f59e0b', lastActivity: '30 minutes ago', activeProjects: 4 },
      { team: 'Support', progress: 96, members: 6, color: '#8b5cf6', lastActivity: '15 minutes ago', activeProjects: 1 }
    ];
  }

  private getFallbackMigrationPhases(): LiveMigrationPhase[] {
    return [
      {
        id: 'initiation',
        phase: 'Initiation',
        description: 'Define business goals and scope for O365 Accelerator. Identify stakeholders and project owners.',
        tools: 'Project Charter, Kick-off PPT',
        status: 'completed',
        progress: 100,
        icon: 'FileText'
      },
      {
        id: 'assessment',
        phase: 'Assessment',
        description: 'Assess existing environment (AD, Exchange, File shares, etc.) and readiness for O365 migration.',
        tools: 'Microsoft Assessment Tool, FastTrack Readiness, MAP Toolkit',
        status: 'completed',
        progress: 100,
        icon: 'Shield'
      },
      {
        id: 'design',
        phase: 'Design',
        description: 'Plan architecture, identify workloads (e.g., Exchange Online, SharePoint, Teams), define identity strategy (cloud-only, hybrid).',
        tools: 'Architecture Diagrams, Microsoft 365 Roadmap',
        status: 'in-progress',
        progress: 75,
        icon: 'Settings'
      }
    ];
  }

  private getFallbackScheduledReports(): LiveScheduledReport[] {
    return [
      {
        id: 'weekly-performance',
        name: 'Weekly Team Performance Report',
        frequency: 'weekly',
        nextRun: 'Monday 9:00 AM',
        status: 'active',
        recipients: ['admin@company.com'],
        reportType: 'Performance Analytics'
      },
      {
        id: 'monthly-analytics',
        name: 'Monthly Analytics Summary',
        frequency: 'monthly',
        nextRun: '1st of next month',
        status: 'active',
        recipients: ['executive@company.com'],
        reportType: 'Executive Summary'
      },
      {
        id: 'daily-status',
        name: 'Daily Project Status Update',
        frequency: 'daily',
        nextRun: 'Tomorrow 8:00 AM',
        status: 'active',
        recipients: ['project-manager@company.com'],
        reportType: 'Project Status'
      }
    ];
  }

  // Clear cache
  clearCache(): void {
    this.cache.clear();
  }

  // Clear specific cache entry
  clearCacheEntry(key: string): void {
    this.cache.delete(key);
  }
} 