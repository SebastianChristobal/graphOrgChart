import {
    Stack,
    FontIcon,
    Text,
    IStackTokens,
    PersonaSize,
    Link,
  } from  "office-ui-fabric-react"
  import * as React from "react";
  import { Person } from "../Person/Person";
  import { IExpandedCardProps } from "./IExpandedCardProps";
  import { useHoverCardStyles } from "./useHoverCardStyles";
  import {
    useGetUserProperties,
    manpingUserProperties,
  } from "../../hooks/useGetUserProperties";
  import { IUserInfo } from "../../models/IUserInfo";
import { OrgChartContext } from "../../context/OrgChartContext";
  export const ExpandedCard: React.FunctionComponent<IExpandedCardProps> = (
    props: IExpandedCardProps
  ) => {
    const { user } = props;
    const { expandedCardStackStyle, hoverCardStyles } = useHoverCardStyles();
  
    const stackFieldTokens: IStackTokens = {
      childrenGap: 15,
  
    };
  /*eslint-disable*/
    const { getUserProfile } = useGetUserProperties();
    const [manager, setManager] = React.useState<IUserInfo>();
    const {extensionAttributeValue} = React.useContext(OrgChartContext);
    React.useEffect(() => {
      if (!user.manager) {
        return;
      }
  
      (async () => {
        const { currentUserProfile } = await getUserProfile(user.manager);
        const wManager: IUserInfo = await manpingUserProperties(
          currentUserProfile
        );
        setManager(wManager);
      })();
    }, [getUserProfile, user?.manager]);
  
    return (
      <>
        <Stack tokens={{ childrenGap: 10 }} styles={expandedCardStackStyle}>
          <Text variant="medium" style={{ fontWeight: 600 }}>
            Contact
          </Text>
          {
            user?.
            mail && (
              <>
              <Stack
              horizontal
              horizontalAlign="start"
              verticalAlign="center"
              styles={{root:{padding: 5}}}
              tokens={stackFieldTokens}
            >
              <FontIcon iconName="mail" className={hoverCardStyles.iconStyles} />
              <Link href={`MAILTO:${user.mail}`}   target="_blank"
                    data-interception="off">
                <Text variant="smallPlus">{user.mail}</Text>
              </Link>
            </Stack>
            <div className={hoverCardStyles.separatorHorizontal}></div>
            </>
            )
          }
  
          {user?.workPhone && (
            <>
              <Stack
                horizontal
                horizontalAlign="start"
                verticalAlign="center"
                tokens={stackFieldTokens}
                styles={{root:{padding: 5}}}
              >
                <FontIcon
                  iconName="Phone"
                  className={hoverCardStyles.iconStyles}
                />
                <Link href={`CALLTO:${user.workPhone}`}   target="_blank"
                  data-interception="off">
                  <Text variant="smallPlus">{user.workPhone}</Text>
                </Link>
              </Stack>
              <div className={hoverCardStyles.separatorHorizontal}></div>
            </>
          )}
          {user?.location && (
            <>
              <Stack
                horizontal
                horizontalAlign="start"
                verticalAlign="center"
                tokens={stackFieldTokens}
                styles={{root:{padding: 5}}}
              >
                <FontIcon
                  iconName="MapPin"
                  className={hoverCardStyles.iconStyles}
                />
                <Link
                  href={`https://www.bing.com/maps?q=${encodeURIComponent(
                    user.location
                  )}`}
                  target="_blank"
                  data-interception="off"
                >
                  <Text variant="smallPlus">{user.location}</Text>
                </Link>
              </Stack>
            </>
          )}
  
          {manager && (
            <>
              <Text
                variant="medium"
                style={{ fontWeight: 600, marginTop: 25, marginBottom: 10 }}
              >
                Reports to
              </Text>
              <Person
                userEmail={manager.mail}
                size={PersonaSize.size48}
                pictureUrl={manager.pictureUrl}
                text={manager.displayName}
                secondaryText={manager.title}
                optionaltext={manager.office}
                tertiaryText={user?.onPremisesExtensionAttributes?.[extensionAttributeValue]}
              ></Person>
            </>
          )}
        </Stack>
      </>
    );
  };
  