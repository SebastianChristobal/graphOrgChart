import * as React from "react";
import { IOrgChartProps } from "./IOrgChartProps";
import { IOrgChartState } from "./IOrgChartState";
import { OrgChartReducer } from "./OrgChartReducer";
import {
  useGetUserProperties,
  manpingUserProperties,
} from "../../hooks/useGetUserProperties";
import { IStackStyles, Stack } from "office-ui-fabric-react/lib/Stack";
import { PersonCard } from "../PersonCard/PersonCard";
import { IUserInfo } from "../../models/IUserInfo";
import { EOrgChartTypes } from "./EOrgChartTypes";
import {
  MessageBar,
  MessageBarType,
  Overlay,
  Spinner,
  SpinnerSize,
  Text,
} from "office-ui-fabric-react";



import { getGUID } from "@pnp/common";
import { useOrgChartStyles } from "./useOrgChartStyles";
import { Placeholder } from "@pnp/spfx-controls-react/lib/controls/placeholder";


const initialState: IOrgChartState = {
 isLoading: true,
  renderDirectReports: [],
  renderManagers: [],
  error: undefined,
  currentUser: undefined,
};

const titleStyle: IStackStyles = {
  root: {
    paddingBottom: 40,
  },
};
/*eslint-disable*/
export const OrgChart: React.FunctionComponent<IOrgChartProps> = (props: IOrgChartProps) => {
  const { getUserProfile } =  useGetUserProperties();
  const [state, dispatch] = React.useReducer(OrgChartReducer, initialState);
  const { orgChartClasses } = useOrgChartStyles();
  
  const {
    renderManagers,
    renderDirectReports,
    currentUser,
    isLoading,
    error,
  }: IOrgChartState = state;

  const {
    context,
    showAllManagers,
    startFromUser,
    showActionsBar,
    title,
    filterByCompanyName
  }: IOrgChartProps = props;

  const startFromUserId: string = React.useMemo(
    () => startFromUser && startFromUser[0].login,
    [startFromUser]
    );
  
  const onUserSelected = React.useCallback((selectedUser: IUserInfo) => {
    dispatch({
      type: EOrgChartTypes.SET_CURRENT_USER,
      payload: selectedUser,
    });
  }, []);
  const loadOrgChart = React.useCallback(

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    async (selectedUser: string): Promise<any> => {

      const wRenderManagers: JSX.Element[] = [];
      const wRenderDirectReports: JSX.Element[] = [];
 
      try {
        const result = await getUserProfile(selectedUser, startFromUserId, showAllManagers, filterByCompanyName);
   
        if (result) {
        const { managersList, reportsLists } = result;
        if (managersList) {
          for (const managerInfo of managersList) {
          
            wRenderManagers.push(
              <>
                <PersonCard
                  key={getGUID()}
                  userInfo={managerInfo}
                  onUserSelected={onUserSelected}
                  selectedUser={currentUser}
                  showActionsBar={showActionsBar}
                ></PersonCard>
                <div
                  key={getGUID()}
                  className={orgChartClasses.separatorVertical}
                ></div>
              </>
            );
          }

          for (const directReport of reportsLists) {
        
            wRenderDirectReports.push(
              <>
                <PersonCard
                  key={getGUID()}
                  userInfo={directReport}
                  onUserSelected={onUserSelected}
                  selectedUser={currentUser}
                  showActionsBar={showActionsBar}
                ></PersonCard>
              </>
            );
          }
        }

        dispatch({
          type: EOrgChartTypes.SET_HAS_ERROR,
          payload: { hasError: false, errorMessage: "" },
        });
      }
      } catch (error) {
        console.log(error.message)
        dispatch({
          type: EOrgChartTypes.SET_IS_LOADING,
          payload: false,
        });
        dispatch({
          type: EOrgChartTypes.SET_HAS_ERROR,
          payload: {
            hasError: true,
            errorMessage: "error",
          },
        });
      }

      return { wRenderDirectReports, wRenderManagers };
    },
    [
      getUserProfile,
      startFromUserId,
      showAllManagers,
      onUserSelected,
      currentUser,
      showActionsBar,
      orgChartClasses.separatorVertical,
    ]
  );
  /*eslint-disable*/
  React.useEffect(() => {
    void fetchData();
  }, [startFromUserId]);
  async function fetchData() {
    
    await loadOrgChart(startFromUserId);
   }
  
  React.useEffect(() => {
    (async () => {
      try {
        const adress = context.pageContext.web.absoluteUrl.replace(context.pageContext.web.serverRelativeUrl, "")
        
        if (startFromUserId === undefined)  return;
        if (startFromUserId === ''){
         
          dispatch({
            type: EOrgChartTypes.SET_IS_LOADING,
            payload: false,
          });
         
          dispatch({
            type: EOrgChartTypes.SET_HAS_ERROR,
            payload: {
              hasError: true,
              errorMessage: "User don't have email defined",
            },
          });
          return;
        }
        const { currentUserProfile } = await getUserProfile(startFromUserId);
        const wCurrentUser: IUserInfo = await manpingUserProperties(
          currentUserProfile,adress
        );
        dispatch({
          type: EOrgChartTypes.SET_CURRENT_USER,
          payload: wCurrentUser,
        });
        dispatch({
          type: EOrgChartTypes.SET_HAS_ERROR,
          payload: { hasError: false, errorMessage: "" },
        });
      } catch (error) {
     
        dispatch({
          type: EOrgChartTypes.SET_IS_LOADING,
          payload: false,
        });
        dispatch({
          type: EOrgChartTypes.SET_HAS_ERROR,
          payload: {
            hasError: true,
            errorMessage: "error",
          },
        });
      }
    })();
  }, [getUserProfile, startFromUserId]);

  React.useEffect(() => {
    (async () => {
      if (!currentUser) return;
      dispatch({
        type: EOrgChartTypes.SET_IS_LOADING,
        payload: true,
      });
      console.log(currentUser)
      const { wRenderDirectReports, wRenderManagers } = await loadOrgChart(
        currentUser.loginName
      );
console.log(currentUser)
      dispatch({
        type: EOrgChartTypes.SET_RENDER_MANAGERS,
        payload: wRenderManagers,
      });
      dispatch({
        type: EOrgChartTypes.SET_RENDER_DIRECT_REPORTS,
        payload: wRenderDirectReports,
      });
      dispatch({
        type: EOrgChartTypes.SET_IS_LOADING,
        payload: false,
      });
    })();
  }, [currentUser, loadOrgChart]);

  if (!startFromUser) {
    return (
      <Placeholder
        iconName="Edit"
        iconText="Configure your Organization Chart Web Part"
        description={"Please configure web part"}
        buttonLabel="Configure"
        onConfigure={context.propertyPane.open}
      />
    );
  }

  if (isLoading) {
    return (
      <Overlay style={{ height: "100%", position: "fixed" }}>
        <Stack style={{ height: "100%" }} verticalAlign="center">
          <Spinner
            styles={{ root: { zIndex: 9999 } }}
            size={SpinnerSize.large}
            label={"loading Organization Chart..."}
            labelPosition={"bottom"}
          ></Spinner>
        </Stack>
      </Overlay>
    );
  }

  if (error && error.hasError) {
    return (
      <Stack
        horizontal
        horizontalAlign="center"
        styles={{root:{padding: 20}}}
        tokens={{ childrenGap: 10 }}
      >
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {error.errorMessage}
        </MessageBar>
      </Stack>
    );
  }

  return (
    <>
      <Stack  styles={{root:{padding: 20}}} >
        <Stack horizontal horizontalAlign="center" styles={titleStyle}>
          <Text variant="xLarge" block>
            {title}
          </Text>
        </Stack>
        <Stack horizontalAlign="center" verticalAlign="center">
          {renderManagers}
          <PersonCard
            key={getGUID()}
            userInfo={currentUser}
            onUserSelected={onUserSelected}
            selectedUser={currentUser }
            showActionsBar={showActionsBar}
          ></PersonCard>
          {renderDirectReports.length && (
            <>
              <div className={orgChartClasses.separatorVertical}></div>
              <div className={orgChartClasses.separatorHorizontal}></div>
            </>
          )}
        </Stack>
        <Stack
          horizontal
          inner={{style:{justifyContent:"center", flexFlow:"wrap", gap:"30px"}}}
          horizontalAlign="center"
          styles={{root:{paddingLeft: 0, paddingRight: 0, paddingBottom: 0, paddingTop:10} }}
          tokens={{ childrenGap: 15 }}
          wrap
        >
          
          {renderDirectReports}
        </Stack>
      </Stack>
    </>
  );
};
