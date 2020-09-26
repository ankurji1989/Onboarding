import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITempTrainingProps {
  context: WebPartContext;
  displayMode: DisplayMode,
  configured: boolean,
  tempTrainingUserList: string;
}
