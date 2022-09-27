import * as React from 'react';
import styles from './SafetyHub.module.scss';
import { ISafetyHubProps } from './ISafetyHubProps';
import dataService from "../../../Common/DataService";
import CONSTANTS from "../../../Common/Constants";
import { escape } from '@microsoft/sp-lodash-subset';
import 'office-ui-fabric-react/dist/css/fabric.css';
//import { Dropdown, IDropdownStyles, IDropdownOption, Label } from 'office-ui-fabric-react';
//import { Dropdown, IDropdownStyles, IDropdownOption, Label, ComboBox, IComboBox, IComboBoxOption, Stack, IStackTokens } from '@fluentui/react';
import * as _ from 'lodash';
import { RxJsEventEmitter } from '../../RxJsEventEmitter/RxJsEventEmitter';
import { Dropdown, IDropdownStyles, IDropdownOption, Label, ComboBox, IComboBox, IComboBoxOption, Stack, IStackTokens } from 'office-ui-fabric-react';
import BlockUi from 'react-block-ui';
import 'react-block-ui/style.css';

let RegionOptions: IDropdownOption[] = [];
let ProgramTypeOptions: IDropdownOption[] = [];
const outerStackTokens: IStackTokens = { childrenGap: 5, padding: 10 };

export interface ISafetyHubState {
  regionOptions: IDropdownOption[];
  programTypeOptions: IDropdownOption[];
  selectedRegionValue: any;
  selectedProgramTypeValue: any;
  programTypeDropDownError: boolean;
  regionDropDownError: boolean;
  blocking: boolean;
}

export interface IEventData {
  sharedRegion: any;
  sharedProgramType: any;
  sharedLeftNavigation: string;
  sharedSiteLocationMetdataAvailable: boolean;
}

const commonService = new dataService();
export default class SafetyHub extends React.Component<ISafetyHubProps, ISafetyHubState> {
  private readonly eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();
  constructor(props: ISafetyHubProps, state: ISafetyHubState) {
    super(props);
    this.state = ({
      regionOptions: [],
      programTypeOptions: [],
      selectedRegionValue: [],
      selectedProgramTypeValue: [],
      programTypeDropDownError: false,
      regionDropDownError: false,
      blocking: false
    });
  }

  public async componentDidMount() {
    this.LoadRegion();
    this.LoadProgramType();

    // this.setState({
    //   blocking: true
    // });

    // setTimeout(() => {
    //   this.setState({ blocking: false });
    // }, 5000);
  }

  private LoadRegion = (): void => {
    RegionOptions = [];
    commonService.GetRegion(CONSTANTS.SITE_COLUMN_NAME.REGION_COLS).then((RegionChoices: any) => {
      RegionOptions.push({
        key: "Select",
        text: "Select"
      });

      RegionChoices.Choices.forEach((RegionItem: any, index: number) => {
        RegionOptions.push({
          key: (index + 1),
          text: RegionItem
        });
      });
      let queryStringParameters = new URLSearchParams(window.location.search);

      let selectedRegion: any[] = [];
      if (queryStringParameters.get("rgn")) {
        selectedRegion = _.filter(RegionOptions, (p) => {
          return p.key == queryStringParameters.get("rgn");
        });
      }

      if (selectedRegion.length > 0) {
        this.setState({
          regionOptions: RegionOptions,
          selectedRegionValue: { key: selectedRegion[0].key, text: selectedRegion[0].text }
        });
      } else {
        this.setState({
          regionOptions: RegionOptions,
          selectedRegionValue: { key: "Select", text: "Select" }
        });
      }

    });
  }

  private LoadProgramType = (): void => {
    ProgramTypeOptions = [];
    commonService.GetProgramType(CONSTANTS.SITE_COLUMN_NAME.PROGRAM_TYPE_COLS).then((ProgramTypeChoices: any) => {
      ProgramTypeOptions.push({
        key: "Select",
        text: "Select"
      });
      ProgramTypeChoices.Choices.forEach((ProgramTypeItem: any, index: number) => {
        ProgramTypeOptions.push({
          key: (index + 1),
          text: ProgramTypeItem
        });
      });

      let queryStringParameters = new URLSearchParams(window.location.search);

      let selectedProgramType: any[] = [];
      if (queryStringParameters.get("pty")) {

        selectedProgramType = _.filter(ProgramTypeOptions, (p) => {
          return p.key == queryStringParameters.get("pty");
        });
      }
      if (selectedProgramType.length > 0) {
        this.setState({
          programTypeOptions: ProgramTypeOptions,
          selectedProgramTypeValue: { key: selectedProgramType[0].key, text: selectedProgramType[0].text }

        });
      } else {
        this.setState({
          programTypeOptions: ProgramTypeOptions,
          selectedProgramTypeValue: { key: "Select", text: "Select" }
        });
      }

    });
  }

  public onRegionDropdownChange = (e: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    if (item.text != "") {
      this.setState({
        selectedRegionValue: item
      });
      this.sendData(item, this.state.selectedProgramTypeValue);
    }
  }

  public onProgramTypeDropdownChange = (e: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    if (item.text != "") {
      this.setState({
        selectedProgramTypeValue: item
      });
      this.sendData(this.state.selectedRegionValue, item);
    }

  }

  private sendData(selectedRegion: any, selectedProgramType: any): void {
    var eventBody = {
      sharedRegion: selectedRegion,
      sharedProgramType: selectedProgramType,
      sharedLeftNavigation: "",
      sharedSiteLocationMetdataAvailable: false
    } as IEventData;

    this.eventEmitter.emit("shareData", eventBody);
  }

  public onProgramTypeChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    if (option != undefined) {
      this.setState({
        ...this.state,
        selectedProgramTypeValue: option,
        programTypeDropDownError: false
      });
      this.sendData(this.state.selectedRegionValue, option);
    } else {
      this.setState({
        ...this.state,
        selectedProgramTypeValue: { key: "Select", text: "Select" },
        programTypeDropDownError: true
      });
    }
  }

  public onRegionChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    if (option != undefined) {
      this.setState({
        ...this.state,
        selectedRegionValue: option,
        regionDropDownError: false
      });
      this.sendData(option, this.state.selectedProgramTypeValue);
    } else {
      this.setState({
        ...this.state,
        selectedRegionValue: { key: "Select", text: "Select" },
        regionDropDownError: true
      });
    }
  }


  public render(): React.ReactElement<ISafetyHubProps> {

    return (
      // <BlockUi blocking={this.state.blocking} message="Loading Content...Please Wait" keepInView={true}>
      <div className={styles.safetyHub}>
        
        <Stack tokens={outerStackTokens}>

          <Stack.Item grow className="SafetyHubFilter">
            <div className={styles.regionDiv}>

              <ComboBox
                label="REGION"
                allowFreeform={true}
                autoComplete={'on'}
                options={this.state.regionOptions}
                onChange={this.onRegionChange}
                selectedKey={this.state.selectedRegionValue ? this.state.selectedRegionValue.key : "Select"}
                errorMessage={this.state.regionDropDownError == true ? "Please select valid region." : ""}
              />
            </div>
            <div className={styles.programTypeDiv}>

              <ComboBox
                label="PROGRAM TYPE"
                allowFreeform={true}
                autoComplete={'on'}
                options={this.state.programTypeOptions}
                onChange={this.onProgramTypeChange}
                selectedKey={this.state.selectedProgramTypeValue ? this.state.selectedProgramTypeValue.key : "Select"}
                errorMessage={this.state.programTypeDropDownError == true ? "Please select valid program type." : ""}
              />
            </div>



          </Stack.Item>
        </Stack>
       
      </div>
      // </BlockUi>
    );
  }

}
