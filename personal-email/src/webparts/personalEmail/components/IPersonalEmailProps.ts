import { DisplayMode } from "@microsoft/sp-core-library"
import { MSGraphClient } from "@microsoft/sp-http"
import { IReadonlyTheme } from "@microsoft/sp-component-base"
import { IPersonalEmailWebPartProps } from "../PersonalEmailWebPart"

export interface IPersonalEmailProps extends IPersonalEmailWebPartProps {
  displayMode: DisplayMode;
  graphClient: MSGraphClient;
  themeVariant: IReadonlyTheme | undefined;
  updateProperty: (value: string) => void;
}
