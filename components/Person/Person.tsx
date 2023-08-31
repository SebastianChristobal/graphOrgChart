import * as React from "react";
import {
  IPersonaSharedProps,
  Persona,
  PersonaSize,
} from "office-ui-fabric-react/lib/Persona";
import { Text } from "office-ui-fabric-react/lib/Text";
import { IPersonProps } from "./IPersonProps";

export const Person: React.FunctionComponent<IPersonProps> = (
  props: IPersonProps
) => {
  const { text, secondaryText, userEmail, size, tertiaryText , pictureUrl, optionaltext} = props;
  const personProps: IPersonaSharedProps = React.useMemo(() => {
    return {
       imageUrl: pictureUrl ? `/_layouts/15/userphoto.aspx?size=M&accountname=${userEmail}` : undefined,
      text: text,
      secondaryText: secondaryText,
      tertiaryText: tertiaryText,
      optionalText: optionaltext
    };
  }, [pictureUrl, userEmail, text, secondaryText, tertiaryText]);

  const _onRenderPrimaryText = React.useCallback(() => {
    return (
      <>
        <Text
          title={text}
          variant="mediumPlus"
          block
          nowrap
          styles={{ root: { fontWeight: 600 } }}
        >
          {text}
        </Text>
      </>
    );
  }, [text]);

  const _onRenderSecondaryText = React.useCallback(() => {
    return (
      <>
        <Text
          title={secondaryText}
          variant="smallPlus"
          block
          nowrap
          styles={{ root: { fontWeight: 400 } }}
        >
          {secondaryText}
        </Text>
      </>
    );
  }, [secondaryText]);

  const _onRenderTertiaryText = React.useCallback(() => {
    return (
      <>
        <Text
          title={tertiaryText}
          variant="smallPlus"
          block
          nowrap
          styles={{ root: { fontWeight: 400 } }}
        >
          {tertiaryText}
        </Text>
      </>
    );
  }, [tertiaryText]);
  const _onRenderOptionalText = React.useCallback(() => {
    return (
      <>
        <Text
          title={optionaltext}
          variant="smallPlus"
          block
          nowrap
          styles={{ root: { fontWeight: 400 } }}
        >
          {optionaltext}
        </Text>
      </>
    );
  }, [tertiaryText]);
  
  return (
    <>
      <Persona
        {...personProps}
        size={size || PersonaSize.size40}
        onRenderPrimaryText={_onRenderPrimaryText}
        onRenderSecondaryText={_onRenderSecondaryText}
        onRenderTertiaryText={_onRenderTertiaryText}
        onRenderOptionalText={_onRenderOptionalText}
        styles={{
          secondaryText: { maxWidth: 230 },
          primaryText: { maxWidth: 230 },
          tertiaryText: {display: "block",maxWidth: 230},
          optionalText:{display: "block",maxWidth: 230}
        }}
      />
    </>
  );
};
