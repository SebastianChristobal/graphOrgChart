import { PersonaSize } from "office-ui-fabric-react/lib/Persona";
export interface IPersonProps {
  userEmail: string;
  text: string;
  secondaryText: string;
  tertiaryText?: string;
  optionaltext?: string;
  pictureUrl?:string;
  size?: PersonaSize;
}
