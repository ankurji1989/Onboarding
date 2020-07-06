import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IOnboardingCheckListProps {
  context: WebPartContext;
  checkList: string;
  onboardingList: string;
  registrationList: string;
}
