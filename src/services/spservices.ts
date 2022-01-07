import { sp } from "@pnp/sp";
import "@pnp/sp/taxonomy";
import "@pnp/sp/webs";
import "@pnp/sp/lists"
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/taxonomy";
import { IFieldInfo } from "@pnp/sp/fields";
import { IDropdownOption } from "office-ui-fabric-react";
import { IPickerTerms} from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { IItemAddResult, IItemDeleteParams } from "@pnp/sp/items";
import { IItemUpdateResult } from "@pnp/sp/items";
import { IClase } from "../models/IClase";
import { getGUID } from "@pnp/common";

export interface ISharePointService {
   
    getItems() : Promise<IClase[]>;
    addItems(
        fechaClase: Date,
        profesorClase: any,
        asignaturaClase: IPickerTerms,
        alumnoClase: any,
        periodoClase: string
        ): Promise<IItemAddResult>;
    getField(field: string): Promise<IDropdownOption[]>;
    updateItems(
        id: number,
        fechaClase: Date,
        profesorClase: any,
        asignaturaClase: IPickerTerms,
        alumnoClase: any,
        periodoClase: string
        ): Promise<IItemUpdateResult>;
    getItem(id:number): Promise<IClase>;
    deleteItems(id:number): Promise<void>;
    
}

export class SharepointService {
    public async getItems(): Promise<IClase[]>{

        return new Promise<IClase[]>(async resolve => {
                
                const items = await sp.web.lists.getByTitle("Clases").items.
                select("*","ProfesorClase/Title","ProfesorClase/ID","TaxCatchAll/ID","TaxCatchAll/Term","AlumnoClase/Title",
                "AlumnoClase/ID" ).expand("ProfesorClase","TaxCatchAll","AlumnoClase").get();
                console.log(items);

                const data:IClase[] = [];

                items.map(item =>{
                
                const asignaturas = [];
                item.AsignaturaClase.map (n=>{
                    asignaturas.push(n.Label)
                })

                const fecha:Date = new Date(item.FechaClase);
                const fecha1 = this.FormatDate(fecha);

                data.push({
                    ID: item.ID,
                    FechaClase: fecha1,
                    ProfesorClase: item.ProfesorClase.Title,
                    ProfesorClaseKey : item.ProfesorClase.ID,
                    AsignaturaClase: asignaturas,
                    AlumnoClase: item.AlumnoClase.Title,
                    AlumnoClaseKey: item.AlumnoClase.ID,
                    PeriodoClase: item.PeriodoClase,
                })
                });

            resolve(data);
        });
    }

    public async addItems(
        fechaClase: Date,
        profesorClase: any,
        asignaturaClase: IPickerTerms,
        alumnoClase: any,
        periodoClase: string
        ): Promise<IItemAddResult>{

        return new Promise<IItemAddResult>(async resolve => {
           
            const list = await sp.web.lists.getByTitle('Clases');
           
            const field = await list.fields.getByTitle('AsignaturaClase_0').get();  
           
            
            let termsString: string = '';
            asignaturaClase.forEach(term => {
            termsString += `-1;#${term.name}|${(term.key)};#`;
            })

            const iar: IItemAddResult = await sp.web.lists.getByTitle("Clases").items.add({
                Title: getGUID(),
                FechaClase: fechaClase,
                ProfesorClaseId: profesorClase,
                ca0ef13dc920495a8398dd8a40d83852: termsString,
                AlumnoClaseId: alumnoClase,
                PeriodoClase: periodoClase,        
              });
              resolve(iar);
            });
    }

    public async getField(field: string): Promise<IDropdownOption[]> {

        return new Promise<IDropdownOption[]>(async resolve => {
                
                const Field: IFieldInfo = await sp.web.lists.getByTitle("Clases").fields.getByTitle(field)();
                const FieldChoices: [] = Field["Choices"];
                const data:IDropdownOption[] = [];
                FieldChoices.map(item =>{
                data.push({
                    key: item,
                    text: item
                })
                });
            resolve(data);
        });
      }


      public async updateItems(
        id: number,
        fechaClase: Date,
        profesorClase: any,
        asignaturaClase: IPickerTerms,
        alumnoClase: any,
        periodoClase: string
        ): Promise<IItemUpdateResult>{

        return new Promise<IItemUpdateResult>(async resolve => {
            
            let termsString: string = '';
            asignaturaClase.forEach(term => {
            termsString += `-1;#${term.name}|${(term.key)};#`;
            })

            const iur: IItemUpdateResult = await sp.web.lists.getByTitle("Clases").items.getById(id).update({
                Title: getGUID(),
                FechaClase: fechaClase,
                ProfesorClaseId: profesorClase,
                ca0ef13dc920495a8398dd8a40d83852: termsString,
                AlumnoClaseId: alumnoClase,
                PeriodoClase: periodoClase,        
              });
              resolve(iur);
            });
    }

    public async getItem(id:number): Promise<IClase>{

        return new Promise<IClase>(async resolve => {
                
                const item = await sp.web.lists.getByTitle("Clases").items.getById(id).select("*","ProfesorClase/Title","ProfesorClase/ID","TaxCatchAll/ID","TaxCatchAll/Term","AlumnoClase/Title",
                "AlumnoClase/ID" ).expand("ProfesorClase","TaxCatchAll","AlumnoClase").get();

                const data: IClase = {
                    ID: item.ID,
                    FechaClase: item.FechaClase,
                    ProfesorClase: item.ProfesorClase.Title,
                    ProfesorClaseKey: item.ProfesorClase.ID,
                    AsignaturaClase: item.AsignaturaClase,
                    AlumnoClase: item.AlumnoClase.Title,
                    AlumnoClaseKey: item.AlumnoClase.ID,
                    PeriodoClase: item.PeriodoClase,
                };

            resolve(data);
        });
    }

    public async deleteItems(id:number): Promise<void>{

        return new Promise<void>(async resolve => {
                
                const idp : void = await sp.web.lists.getByTitle("Clases").items.getById(id).delete();
                
               
            resolve(idp);
        });
    }
    
    public FormatDate = (date): string => {
        var date1 = new Date(date);
        var year = date1.getFullYear();
        var month = (1 + date1.getMonth()).toString();
        month = month.length > 1 ? month : '0' + month;
        var day = date1.getDate().toString();
        day = day.length > 1 ? day : '0' + day;
        return day + '/' + month + '/' + year;
      };


}