/* tslint:disable */
import * as React from "react";
import {
  DocumentCard,
  DocumentCardActions,
  DocumentCardDetails,
  IDocumentCard,
} from "office-ui-fabric-react/lib/DocumentCard";

import {
  Stack,
  IButtonProps,
  HoverCard,
  HoverCardType,
  IExpandingCardProps,
  PersonaSize,
  DirectionalHint,
} from "office-ui-fabric-react";
import { Person } from "../Person/Person";
import { IUserInfo } from "../../models/IUserInfo";
import { ExpandedCard, CompactCard } from "../HoverCard";

import { IPersonCardProps } from "./IPersonCardProps";
import { usePersonaCardStyles } from "./usePersonCardStyles";
import { OrgChartContext } from "../../context/OrgChartContext";

const currentTheme = window.__themeState__.theme;

export const PersonCard: React.FunctionComponent<IPersonCardProps> = (props: IPersonCardProps) => {
  const {extensionAttributeValue} = React.useContext(OrgChartContext);
  const { userInfo, onUserSelected, showActionsBar } = props;

  const documentCardRef = React.useRef<IDocumentCard>(undefined);
  const {
    personaCardStyles,
    documentCardActionStyles,
    buttonStyles,
    stackPersonaStyles,
  } = usePersonaCardStyles();

  const documentCardActions: IButtonProps[] = React.useMemo(() => {
    const cardActions: IButtonProps[] = [];
    cardActions.push({
      iconProps: { iconName: "Chat" },
      title: "Chat",
      styles: buttonStyles,
      onClick: (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        window.open(
          `https://teams.microsoft.com/l/chat/0/0?users=${userInfo.mail}&message=Hi ${userInfo.displayName} `,
          "_blank"
        );
      },
    });

    if (userInfo?.mail) {
      cardActions.push({
        iconProps: { iconName: "Mail" },
        title: "Mail",
        styles: buttonStyles,
        onClick: (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          window.open(`MAILTO:${userInfo.mail}`, "_blank");
        },
      });
    }

    if (userInfo?.workPhone) {
      cardActions.push({
        iconProps: { iconName: "Phone" },
        title: "Call",
        styles: buttonStyles,
        onClick: (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          window.open(`CALLTO:${userInfo.workPhone}`, "_blank");
        },
      });
    }

    if (userInfo?.hasDirectReports) {
      cardActions.push({
        iconProps: { iconName: "Org" },
        title: "View Organization",
        styles: { ...buttonStyles },
      });
    }
    return cardActions;
  }, [
    buttonStyles,
    userInfo?.displayName?.toString(),
    userInfo?.mail,
    userInfo?.hasDirectReports,
    userInfo?.workPhone,
  ]);

  const onRenderCompactCard = React.useCallback(
    (user: IUserInfo): JSX.Element => <CompactCard user={user} />,
    []
  );

  const onRenderExpandedCard = React.useCallback(
    (user: IUserInfo): JSX.Element => <ExpandedCard user={user} />,
    []
  );

  const expandingCardProps: IExpandingCardProps = React.useMemo(() => {
    return {
      onRenderCompactCard: onRenderCompactCard,
      onRenderExpandedCard: onRenderExpandedCard,
      renderData: userInfo,
      directionalHint: DirectionalHint.rightTopEdge,
      styles: {
        expandedCard: { backgroundColor: currentTheme.neutralLighterAlt },
      },
      gapSpace: 5,
      calloutProps: {
        isBeakVisible: false,
      },
    };
  }, [onRenderCompactCard, onRenderExpandedCard, userInfo]);

  return (
    <>
      <DocumentCard
        className={personaCardStyles.tile}
        style={{marginLeft:0, marginRight: 0, marginTop: 7.5, marginBottom: 7.5, maxWidth:"342px"}}
        componentRef={documentCardRef}
        onClick={() => {
          //  documentCardRef.current.focus();
          if (userInfo.hasDirectReports) {
            onUserSelected(userInfo);
            //
          }
        }}
      >
        <HoverCard
          expandingCardProps={expandingCardProps}
          type={HoverCardType.expanding}
        
        >
          <DocumentCardDetails styles={{root:{height:80}}}>
            <Stack
              horizontal
              horizontalAlign="space-between"
              styles={stackPersonaStyles}
            >
              <Person
                text={userInfo?.displayName}
                secondaryText={userInfo?.title}
                tertiaryText={userInfo?.onPremisesExtensionAttributes?.[extensionAttributeValue]}
                optionaltext={userInfo?.office}
                userEmail={userInfo?.mail}
                pictureUrl={userInfo?.pictureUrl}
                size={PersonaSize.size40}
              />
            </Stack>
          </DocumentCardDetails>
        </HoverCard>
        {showActionsBar && (
          <DocumentCardActions
            actions={documentCardActions}
            styles={documentCardActionStyles}
          />
        )}
      </DocumentCard>
    </>
  );
};
