import * as React from 'react';
import styles from './PersonalEmail.module.scss';
import * as strings from 'PersonalEmailWebPartStrings';
import { IMessages, IMessage, IPersonalEmailProps, IPersonalEmailState } from '.';

import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';

import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { List } from 'office-ui-fabric-react/lib/components/List';
import { Link } from 'office-ui-fabric-react/lib/components/Link';
import { IIconProps } from 'office-ui-fabric-react/lib/components/Icon';
import { ActionButton } from 'office-ui-fabric-react/lib/components/Button';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export default class PersonalEmail extends React.Component<IPersonalEmailProps, IPersonalEmailState> {

  protected readonly Outlook: string = "https://outlook.office.com/owa/";
  protected readonly OutlookNewEmail: string = "https://outlook.office.com/mail/deeplink/compose";

  constructor(props: IPersonalEmailProps) {
    super(props);
    this.state = {
      loading: false,
      messages: [],
      error: undefined,
    };
  }

  private addIcon: IIconProps = { iconName: 'Add' };
  private viewAll: IIconProps = { iconName: 'AllApps' };

  private messageLoader(): void {
    if (!this.props.graphClient) {
      return;
    }

    this.setState({
      loading: true,
      messages: [],
      error: null
    })

    this.props.graphClient
      .api('me/mailFolders/Inbox/messages')
      .version('v1.0')
      .select('bodyPreview,receivedDateTime,from,isRead,subject,webLink')
      .top(this.props.numOfMessages || 5)
      .orderby('receivedDateTime desc')
      .get((err: any, res: IMessages): void => {
        if (err) {
          this.setState({
            error: err.message ? err.message : strings.Error,
            loading: false
          });
          return;
        }

        if (res && res.value && res.value.length > 0) {
          this.setState({
            messages: res.value,
            loading: false
          });
        }

        else {
          this.setState({
            loading: false
          });
        }
      });
  }

  private onRenderCell = (item: IMessage, index: number | undefined): JSX.Element => {
    if (item.isRead) {
      styles.message = styles.message + ' ' + styles.isRead;
    }

    return <Link href={item.webLink} className={styles.message} target='_blank'>
        <div className={styles.from}>
          {item.from.emailAddress.name || item.from.emailAddress.address}
        </div>
        <div className={styles.subject}>
          {item.subject}
        </div>
        <div className={styles.date}>
          {(new Date(item.receivedDateTime).toLocaleDateString())}
        </div>
        <div className={styles.preview}>
          {item.bodyPreview}
        </div>
      </Link>;
  }

  public componentDidMount(): void {
    this.messageLoader();
  }

  public componentDidUpdate(prevProps: IPersonalEmailProps, prevState: IPersonalEmailState): void {
    if (prevProps.numOfMessages !== this.props.numOfMessages) {
      this.messageLoader();
    }
  }

  public render(): React.ReactElement<IPersonalEmailProps> {
    const variantStyles = {
      '--varientBGColor': this.props.themeVariant.semanticColors.bodyBackground,
      '--varientFontColor': this.props.themeVariant.semanticColors.bodyText,
      '--varientBGHovered': this.props.themeVariant.semanticColors.listHeaderBackgroundHovered
    } as React.CSSProperties;

    return (
      <div className={ styles.personalEmail } style={ variantStyles }>
        <WebPartTitle 
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty}
          className={styles.title}
        />
        <ActionButton 
          text={strings.NewEmail}
          iconProps={this.addIcon}
          onClick={this.openOutlookNewEmail}
        />
        <ActionButton
          text={strings.ViewAll}
          iconProps={this.viewAll}
          onClick={this.openOutlook}
        />
        { this.state.loading && <Spinner label={strings.Loading} size={SpinnerSize.large}/> }
        {
          this.state.messages && this.state.messages.length > 0 ? (
            <div>
              <List 
                items={this.state.messages}
                onRenderCell={this.onRenderCell}
                className={styles.list}
              />
            </div>
          ) : (
            !this.state.loading && (
              this.state.error ? 
                <span className={styles.error}>{this.state.error}</span> : 
                <span className={styles.noMessages}>{ strings.NoMessages }</span>
            )
          )
        }
      </div>
    );
  }

  private openOutlookNewEmail = () => {
    window.open(this.OutlookNewEmail, '_blank');
  }

  private openOutlook = () => {
    window.open(this.Outlook, '_blank');
  }
}
