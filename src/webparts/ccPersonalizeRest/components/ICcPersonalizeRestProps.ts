export interface ICcPersonalizeRestProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userLoginName: string;
  userEmail: string;

  // ThirdPartyAPI
  APIResult: any;
}
