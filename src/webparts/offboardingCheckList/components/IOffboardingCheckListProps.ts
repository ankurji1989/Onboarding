import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IOffboardingCheckListProps {
  context: WebPartContext;
  checkList: string;
  onboardingList: string;
}
