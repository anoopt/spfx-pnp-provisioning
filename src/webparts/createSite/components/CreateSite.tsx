import * as React from 'react';
import styles from './CreateSite.module.scss';
import { DefaultButton, PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { ICreateSiteProps } from './ICreateSiteProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ISubSite } from '../../../interfaces/ISubSite';
import { ISubSiteService, SubSiteService } from '../../../services';
import { ICreateSiteState } from './ICreateSiteState';
import { OfficeUiFabricPeoplePicker } from "spfx-office-ui-fabric-people-picker/lib/OfficeUiFabricPeoplePicker";
import { SharePointUserPersona } from 'spfx-office-ui-fabric-people-picker/lib';
import TaxonomyPicker from "react-taxonomypicker";
import "react-taxonomypicker/dist/React.TaxonomyPicker.css";
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { SPComponentLoader } from '@microsoft/sp-loader';

export default class CreateSite extends React.Component<ICreateSiteProps, ICreateSiteState> {

  private _subSiteServiceInstance: ISubSiteService;

  constructor(props: ICreateSiteProps) {
    super(props);
    this.state = {
      loadingScripts: true,
      errors: [],
      status: <span></span>,
      subSiteInContext: {
        Title: "",
        Description: "",
        Url: "",
        GroupName: "",
        GroupOwnerId: 0,
        GroupMemberId: 0,
        Region: {
          Label: "",
          TermGuid: ""
        }
      }
    };
  }

  public componentDidMount(): void {
    this._loadSPJSOMScripts();
  }

  public render(): React.ReactElement<ICreateSiteProps> {
    this._subSiteServiceInstance = this.props.context.serviceScope.consume(SubSiteService.serviceKey);
    return (
      <div className={styles.createSite}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <TextField label='Title' onChanged={this._setTitle.bind(this)} />
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <TextField label='Description' multiline rows={4} onChanged={this._setDescription.bind(this)} />
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <TextField label='Url' onChanged={this._setUrl.bind(this)} />
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <TextField label='Group Name' onChanged={this._setGroupName.bind(this)} />
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <Label>Group Owner</Label>
              <OfficeUiFabricPeoplePicker
                description=""
                spHttpClient={this.props.context.spHttpClient}
                siteUrl={this.props.context.pageContext.web.absoluteUrl}
                typePicker="Normal"
                principalTypeUser={true}
                principalTypeSharePointGroup={true}
                principalTypeSecurityGroup={false}
                principalTypeDistributionList={false}
                numberOfItems={1}
                onChange={this._setGroupOwner.bind(this)} />
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <Label>Group Member</Label>
              <OfficeUiFabricPeoplePicker
                description=""
                spHttpClient={this.props.context.spHttpClient}
                siteUrl={this.props.context.pageContext.web.absoluteUrl}
                typePicker="Normal"
                principalTypeUser={true}
                principalTypeSharePointGroup={true}
                principalTypeSecurityGroup={false}
                principalTypeDistributionList={false}
                numberOfItems={1}
                onChange={this._setGroupMember.bind(this)} />
            </div>
          </div>
          {this.state.loadingScripts === false ?
            <div className={styles.row}>
              <div className={styles.column}>
                <Label>Region</Label>
                <TaxonomyPicker
                  name=""
                  displayName=""
                  termSetGuid="59525237-1ba5-4f63-b967-1520f32adb6d"
                  termSetName="Regions"
                  multi={false}

                  onPickerChange={this._setRegion.bind(this)}
                />
              </div>
            </div> : null}
          <div className={styles.row}>
            <div className={styles.column}>
              {this.state.errors}
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <PrimaryButton data-id="btnCreateSubSite"
                title="Create Subsite"
                text="Create Subsite"
                onClick={this._addSubSite.bind(this)}
              />
              <div className={styles.ccStatus}>
                {this.state.status}
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _loadSPJSOMScripts() {
    const siteColUrl = this.props.context.pageContext.site.absoluteUrl;
    try {
      SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/init.js', {
        globalExportsName: '$_global_init'
      })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/MicrosoftAjax.js', {
            globalExportsName: 'Sys'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.Runtime.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.taxonomy.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): void => {
          this.setState({ loadingScripts: false });
        })
        .catch((reason: any) => {
          this.setState({ loadingScripts: false, errors: [...this.state.errors, reason] });
        });
    } catch (error) {
      this.setState({ loadingScripts: false, errors: [...this.state.errors, error] });
    }
  }

  private _setTitle(value: string): void {
    this._setValue("Title", value);
  }

  private _setDescription(value: string): void {
    this._setValue("Description", value);
  }

  private _setUrl(value: string): void {
    this._setValue("Url", value);
  }

  private _setGroupName(value: string): void {
    this._setValue("GroupName", value);
  }

  private _setGroupOwner(items: SharePointUserPersona[]): void {
    this._setValue("GroupOwnerId", items[0].User.Id);
  }

  private _setGroupMember(items: SharePointUserPersona[]): void {
    this._setValue("GroupMemberId", items[0].User.Id);
  }

  private _setRegion = (name, value): void => {
    console.log(value);
    if (value !== null && value !== undefined) {
      let regionLabel: string = value.label.toString();
      let regionTermGuid: string = value.value.toString();
      let subSiteInContext: ISubSite = this.state.subSiteInContext;
      subSiteInContext.Region.Label = regionLabel;
      subSiteInContext.Region.TermGuid = regionTermGuid;
      this.setState({ ...this.state, subSiteInContext });
    }
  }

  private _setValue(field: string, value: any): void {
    let subSiteInContext: ISubSite = this.state.subSiteInContext;
    subSiteInContext[field] = value;
    this.setState({ ...this.state, subSiteInContext });
  }

  private async _addSubSite(): Promise<void> {
    let status: JSX.Element = <Spinner size={SpinnerSize.small} />;
    this.setState({ ...this.state, status });

    let subSiteToAdd: ISubSite = {
      Title: this.state.subSiteInContext.Title,
      Description: this.state.subSiteInContext.Description,
      Url: this.state.subSiteInContext.Url,
      GroupName: this.state.subSiteInContext.GroupName,
      GroupOwnerId: this.state.subSiteInContext.GroupOwnerId,
      GroupMemberId: this.state.subSiteInContext.GroupMemberId,
      Region: this.state.subSiteInContext.Region
    }
    let result: boolean = await this._subSiteServiceInstance.addSubSite(subSiteToAdd);
    console.log(result);
    status = <span></span>;
    this.setState({ ...this.state, status });
  }
}
