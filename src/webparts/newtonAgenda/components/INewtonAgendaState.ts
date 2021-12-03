import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { IDropdownOption } from "office-ui-fabric-react";
import { IClase } from "../../../models/IClase";
import { IViewField } from "@pnp/spfx-controls-react/lib/ListView";

export interface INewtonAgendaState {
  columns: IViewField[];
  items: IClase[];
  periodo: IDropdownOption[];
  fechaClase?: Date;
  profesorClase?: { key: string; name: string } [];
  profesorClaseId?: any;
  asignaturaClase?: IPickerTerms;
  asignaturaClaseP?: IPickerTerms;
  alumnoClase?: { key: string; name: string } [];
  alumnoClaseId?: any;
  periodoClase? : string;
  periodoClaseKey? : string|number;
  visible: boolean;
  selectionId?: number;
  item?: IClase;
  editar : boolean;
  error?: string;
  info?: string;
 
}