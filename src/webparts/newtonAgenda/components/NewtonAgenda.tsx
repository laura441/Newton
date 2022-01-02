import * as React from 'react';
import styles from './NewtonAgenda.module.scss';
import { INewtonAgendaProps } from './INewtonAgendaProps';
import { INewtonAgendaState } from './INewtonAgendaState';
import { ISharePointService, SharepointService } from '../../../services/spservices';
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { IClase } from '../../../models/IClase';
import { Stack, IStackTokens, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
import { PrimaryButton} from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker';
import { ListPicker, ListItemPicker } from "@pnp/spfx-controls-react/lib";
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';
import { DayOfWeek } from '@fluentui/date-time-utilities/lib/dateValues/dateValues';
import { Text, ITextProps } from 'office-ui-fabric-react/lib/Text';
import {MessageBar, MessageBarType} from 'office-ui-fabric-react';
import 'office-ui-fabric-react/dist/css/fabric.css';


let _spService: ISharePointService;

export default class NewtonAgenda extends React.Component<INewtonAgendaProps, INewtonAgendaState> {

  constructor(props: INewtonAgendaProps) {
    super(props);
    
    let spService = new SharepointService();
    _spService = spService;


  this.state = { 
    columns: this.columns(),
    items: [],
    periodo: [],
    visible: false,
    editar: false,
  };

} 

public columns = (): IViewField[]  =>{
  return [ 
    {
      name: "ID",
      displayName: "ID",
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
    },
    {
      name: "FechaClase",
      displayName: "Fecha clase",
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
    },
    {
      name: "ProfesorClase",
      displayName: "Profesor",
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
    },
    {
      name: "AsignaturaClase",
      displayName: "Asignatura/s",
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      render: (item: any) => {
        if(item['AsignaturaClase.0'] && item['AsignaturaClase.1'] && item['AsignaturaClase.2']){
          return <span>{item['AsignaturaClase.0']}<br></br>{item['AsignaturaClase.1']}<br></br>
          {item['AsignaturaClase.2']}<br></br></span>
        }
        else if (item['AsignaturaClase.0'] && item['AsignaturaClase.1']){
          return <span>{item['AsignaturaClase.0']}<br></br>{item['AsignaturaClase.1']}<br></br></span>
        }
        else {
          return <span>{item['AsignaturaClase.0']}<br></br></span>
        }
      }
    },
    {
      name: "AlumnoClase",
      displayName: "Alumno",
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
    },
    {
      name: "PeriodoClase",
      displayName: "Turno",
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
    }

  ]}

  public componentDidMount(): void {
    this._getRequests();
  }

  private async _getRequests(): Promise<void> {

    const items = await _spService.getItems();
    const periodos = await _spService.getField('PeriodoClase');    
    this.setState({items:items,periodo:periodos})
  
}

  

  public render(): React.ReactElement<INewtonAgendaProps> {
    const items = this.state.items;
    const columns = this.state.columns;
    const groupByFields: IGrouping[] = [
      {
        name: "FechaClase", 
        order: GroupOrder.ascending 
      }, {
        name: "PeriodoClase", 
        order: GroupOrder.ascending
      }
    ];
    const today = new Date();
    const todayForm = today.toLocaleDateString();

    return (
 
        <div className={styles.newtonAgenda}>
       { (this.state.visible === false)? 

        <div className={styles.container}>

          <Text variant={'xxLargePlus' as ITextProps['variant']} >
          {"Newton's Agenda"}
          </Text>

          <div className={styles.row}>
              {this.state.info?
                <MessageBar
                        messageBarType={MessageBarType.success}
                        isMultiline={false}
                        dismissButtonAriaLabel="Close"
                        onDismiss={(ev)=>{this.setState({info:null})}}
                >
                 {this.state.info}
                </MessageBar>
                      :
                      <span></span>
              }  
          </div>

          
          <div className={styles.row}>
          <PrimaryButton text="Añadir clase" onClick={this._showForm} className={styles.button1}  />   
          <PrimaryButton text="Editar clase" onClick={this._edClick} className={styles.button1} disabled={(this.state.selectionId)? false:true} />
          <PrimaryButton text="Eliminar clase" onClick={this._deleteClick} className={styles.button1} disabled={(this.state.selectionId)? false:true} />
          </div>
         
          
          <ListView
            items={items}
            compact={false}
            viewFields={columns}
            selectionMode = {SelectionMode.single}
            selection={(item)=>this._getSelection(item)}
            groupByFields={groupByFields}
            showFilter={true}
            defaultFilter={todayForm}
            filterPlaceHolder="Buscar..."
            
          />
        </div>
            
        :
          <div className={styles.container2}>
          <div className= 'ms-Grid' >
          
          <div className={styles.row1}>
          <div className="ms-Grid-row">  
          <div className = "ms-Grid-col ms-sm12 ms-md12 ms-lg12">

          { (this.state.editar === false)?
          <Label className= {styles.title}>Clase creada por: {this.props.context.pageContext.user.displayName}</Label> 
            :
          <Label className= {styles.title}>Clase modificada por: {this.props.context.pageContext.user.displayName}</Label>
          }
         
          </div>
          </div>
          </div>

          <div className={styles.row}>
          <div className="ms-Grid-row">  
          <div className = "ms-Grid-col ms-sm12 ms-md12 ms-lg12">
          {this.state.error?
             
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
                onDismiss={(ev)=>{this.setState({error:null})}}
              >
                {this.state.error}
              </MessageBar>
              :
              <span></span>
            }  
          </div>
          </div> 
          </div>       

          <div className={styles.row}>
          <div className="ms-Grid-row"> 
          <div className = "ms-Grid-col ms-sm12 ms-md12 ms-lg12">
          <DateTimePicker label="Fecha Clase"
                dateConvention={DateConvention.Date}
                value={this.state.editar? (this.state.fechaClase? this.state.fechaClase : new Date (this.state.item.FechaClase)):(this.state.fechaClase? this.state.fechaClase : undefined)}
                minDate={today}
                formatDate= {(date)=>date.toLocaleDateString()}
                onChange={(date)=>this.handleChangeDate(date)}
                firstDayOfWeek={DayOfWeek.Monday}
                 />
          </div>
          </div>
          </div>

          <div className={styles.row}>
          <div className="ms-Grid-row">         
          <div className = "ms-Grid-col ms-sm12 ms-md12 ms-lg6">
          
          <ListItemPicker listId='8301b042-38c2-47b2-a2ee-8a9265c8415b'
                columnInternalName='NombreProfesor'
                keyColumnInternalName='Id'
                itemLimit={1}
                onSelectedItem={(item)=>this.onSelectedItemProfesor(item)}
                context={this.props.context} 
                label= "Profesor"
                defaultSelectedItems= {this.state.editar? (this.state.profesorClase? this.state.profesorClase : [{key:this.state.item.ProfesorClaseKey, name:this.state.item.ProfesorClase}])
                :(this.state.profesorClase? this.state.profesorClase : undefined)}
                />
          
          
          <TaxonomyPicker allowMultipleSelections={true}
                
                termsetNameOrID="AsignaturaClase"
                panelTitle="Seleccione la/s asignatura/s a impartir"
                label="Asignatura/s"
                context={this.props.context}
                onChange={(value)=>this.setAsignatura(value)}
                initialValues = {this.state.editar? (this.state.asignaturaClase? this.state.asignaturaClase : this.state.asignaturaClaseP): 
                  (this.state.asignaturaClase? this.state.asignaturaClase : undefined) }
                 />
          </div>
          <div className = "ms-Grid-col ms-sm12 ms-md12 ms-lg6">
          
          <ListItemPicker listId='e9852d0c-9fd5-41ed-8a83-4495e3004f01'
                columnInternalName='NombreAlumno'
                keyColumnInternalName='Id'
                itemLimit={1}
                onSelectedItem={(item)=>this.onSelectedItemAlumno(item)}
                context={this.props.context} 
                label= "Alumno"
                defaultSelectedItems= {this.state.editar? (this.state.alumnoClase? this.state.alumnoClase: [{key:this.state.item.AlumnoClaseKey, name:this.state.item.AlumnoClase}])
                :(this.state.alumnoClase? this.state.alumnoClase : undefined)}
                />

          <Dropdown
                  label="Turno"
                  onChange={(ev,option)=>this.setPeriodo(option)}
                  options={this.state.periodo}
                  selectedKey= {this.state.editar?(this.state.periodoClase?this.state.periodoClaseKey:this.state.item.PeriodoClase):(this.state.periodoClase?this.state.periodoClaseKey:undefined)}
            />
          
          </div>
          </div>
          </div>

          <div className={styles.row}>
          <div className="ms-Grid-row">         
          <div className = "ms-Grid-col ms-sm12 ms-md12 ms-lg12">
          
          { (this.state.editar === true)?
           <div className={styles.row}>
           
            <PrimaryButton text="Actualizar" onClick={this._editClick} className={styles.button1} disabled={this._disableEdit()}/> 
            <PrimaryButton text="Cancelar" onClick={this._cancelClick} className={styles.button1}/> 
            </div>
          :
            <div className={styles.row}>
            <PrimaryButton text="Guardar" onClick={this._saveClick} className={styles.button1} disabled={this._disableAdd()}/>  
            <PrimaryButton text="Cancelar" onClick={this._cancelClick} className={styles.button1} /> 
            
          </div>
          }
          </div>
          </div>
          </div>
          
         </div>
         </div>
        
        } 
        </div>
    );
  }

  //Método para limpiar el estado
  private cleanState (){
    this.setState({
      fechaClase: null,
      profesorClase: null,
      profesorClaseId: null,
      asignaturaClase: null,
      asignaturaClaseP: null,
      alumnoClase: null,
      alumnoClaseId:null,
      periodoClase: null,
      periodoClaseKey: null,
      visible: false,
      selectionId: null,
      item: null,
      error:null,
      info:null,
    })
  }

  //Método deshabilitar botón guardar
  private _disableAdd ():boolean {
    if (this.state.fechaClase&&this.state.profesorClaseId&&this.state.asignaturaClase&&this.state.alumnoClaseId&&
      this.state.periodoClase&&this.state.asignaturaClase.length!==0){return false}
    else{return true}
  }
  //Método deshabilitar botón actualizar
  private _disableEdit ():boolean {
    if (this.state.asignaturaClase&&this.state.profesorClase&&this.state.alumnoClase){
      if (this.state.asignaturaClase.length!==0&&this.state.profesorClase.length!==0&&this.state.alumnoClase.length!==0)
      {return false}
      else{return true}
    }
    else if (this.state.asignaturaClase&&this.state.profesorClase){
      if (this.state.asignaturaClase.length!==0&&this.state.profesorClase.length!==0)
      {return false}
      else{return true}
    }
    else if (this.state.asignaturaClase&&this.state.alumnoClase){
      if (this.state.asignaturaClase.length!==0&&this.state.alumnoClase.length!==0)
      {return false}
      else{return true}
    }
    else if(this.state.profesorClase&&this.state.alumnoClase){
    if (this.state.profesorClase.length!==0&&this.state.alumnoClase.length!==0)
    {return false}
    else{return true}
    }
    else if(this.state.asignaturaClase){
      if (this.state.asignaturaClase.length!==0)
      {return false}
      else{return true}
    }
    else if (this.state.profesorClase){
      if (this.state.profesorClase.length!==0)
      {return false}
      else{return true}
    }
    else if (this.state.alumnoClase){
      if (this.state.alumnoClase.length!==0)
      {return false}
      else{return true}
    }
    else{
      if (this.state.asignaturaClaseP.length!==0)
      {return false}
      else{return true}
    }
  }

  //Método para guardar el periodo en el estado
  private setPeriodo (value:IDropdownOption){
    const text = value.text;
    const key = value.key;
    this.setState({periodoClase: text, periodoClaseKey: key});
  }
  //Método para guardar la asignatura en el estado
  private setAsignatura(terms : IPickerTerms) {
    this.setState({asignaturaClase: terms});
  }
  //Método para guardar la fecha en el estado
  private handleChangeDate(date:Date){
    this.setState({fechaClase:date});
  }
  //Método al que se llama cuando se cancela el formulario
  private _cancelClick = () => {
    this.cleanState();
    this.setState({editar:false});
  }
  //Método al que se llama cuando se muestra el formulario
  private _showForm = () => {
    this.setState({visible: true});
  }
  //Método al que se llama cuando se guarda una clase
  private  _saveClick = async () =>{
    const fecha = this.state.fechaClase;
    const fecha1 = fecha.setHours(12);
    const fechaClase = new Date (fecha1);
    const profesorClase = this.state.profesorClaseId;
    const asignaturaClase = this.state.asignaturaClase;
    const alumnoClase = this.state.alumnoClaseId;
    const periodoClase = this.state.periodoClase;

    let existeProfe = false;
    let existeAlumno = false;
    let numClase = 0;

    this.state.items.map(item =>{
      const fechaIt = fechaClase.toLocaleDateString() ;
      if (item.FechaClase == fechaIt ){
        if (item.PeriodoClase == periodoClase){
          numClase = numClase + 1 ;
          
          if (item.ProfesorClaseKey == profesorClase){
            existeProfe = true;
          }
 
        }

        if (item.AlumnoClaseKey == alumnoClase) {
        existeAlumno = true;
        }
      }

    })

    if (!existeProfe && !existeAlumno && numClase<4){

      const save = await _spService.addItems(
      fechaClase,
      profesorClase,
      asignaturaClase,
      alumnoClase,
      periodoClase);

      if (save !== null){
        this.cleanState();
        const items = await _spService.getItems();
        this.setState ( {items: items, info: "Se ha guardado correctamente"});
      }else{
        alert("Ha habido un error")
      }

    }else {
      if (existeProfe){ this.setState({error : "Profesor ocupado en ese turno, cambie su opcion"})}
      else if (existeAlumno) {this.setState({error :"Alumno con clase ese día, cambie su opcion"})}
      else if (numClase>=4){this.setState({error :"Turno completo, pruebe con otro o cambie de dia"})}
    }

  }
  //Método para guardar en el estado el profesor
  private onSelectedItemProfesor(data : { key: string; name: string } []) {
    let k : string;
    for (const item of data){
      k = item.key;
    }
    this.setState({profesorClase:data, profesorClaseId:k});
  }
  //Método para guardar en el estado al alumno
  private onSelectedItemAlumno(data : { key: string; name: string } []) {
    let k : string;
    for (const item of data){
      k = item.key;
    }
    this.setState({alumnoClase: data,alumnoClaseId:k});
  }
  //Método para guardar en el estado el item seleccionado
  private _getSelection(items: IClase[]) {
    if (items.length!== 0){
    const itemID = items['0'].ID;
    this.setState({selectionId:itemID});
    }
    else {
    this.setState({selectionId:null})
    }
  }
  //Método al que se llama cuando se clica en editar en la pantalla principal
  private _edClick = async() => {
    const id = this.state.selectionId;
    const item = await _spService.getItem(id);

      if (item !== null){
       
        const pickerTerms : IPickerTerms = [];
        item.AsignaturaClase.map (n=>{
          pickerTerms.push({key:n.TermGuid,name:n.Label,path:n.Label,termSet:'2f456e07-abef-4dac-b19e-402f70cb4afa'})
       });
        this.setState({item:item, visible: true, editar: true, asignaturaClaseP:pickerTerms});

      }else {
        alert("Error cargando item");
      }

  }
  //Método al que se llama cuando se modifica una clase
  private  _editClick = async () =>{
    const id = this.state.item.ID;
    const fecha = this.state.fechaClase?this.state.fechaClase:new Date(this.state.item.FechaClase);
    const fecha1 = fecha.setHours(12);
    const fechaClase = new Date (fecha1);
    const profesorClase = this.state.profesorClaseId?this.state.profesorClaseId:this.state.item.ProfesorClaseKey;
    const asignaturaClase = this.state.asignaturaClase?this.state.asignaturaClase:this.state.asignaturaClaseP;
    const alumnoClase = this.state.alumnoClaseId?this.state.alumnoClaseId:this.state.item.AlumnoClaseKey;
    const periodoClase = this.state.periodoClase?this.state.periodoClase:this.state.item.PeriodoClase;

    let existeProfe = false;
    let existeAlumno = false;
    let numClase = 0;

    const fechaEscogida = fechaClase.toLocaleDateString() ;
    const idEscogido = this.state.item.ID;
      

    this.state.items.map(item =>{

        if (item.ID !== idEscogido) {

            if (item.FechaClase == fechaEscogida ){
              if (item.PeriodoClase == periodoClase){
                numClase = numClase + 1 ;
                
                if (item.ProfesorClaseKey == profesorClase){
                  existeProfe = true;
                }
      
              }

              if (item.AlumnoClaseKey == alumnoClase) {
              existeAlumno = true;
              }
            }

        }

    })

    if (!existeProfe && !existeAlumno && numClase<4){

        const update = await _spService.updateItems(
        id,
        fechaClase,
        profesorClase,
        asignaturaClase,
        alumnoClase,
        periodoClase);

        if (update !== null){
          this.cleanState();
          const items = await _spService.getItems();
          this.setState ( {items: items, editar:false,info:"Se ha editado correctamente"});
        }else{
          alert("Ha habido un error");
        }

  }else {
    if (existeProfe){ this.setState({error :"Profesor existente, cambie su opcion"})}
    else if (existeAlumno) {this.setState({error : "Alumno existente, cambie su opcion"})}
    else if (numClase>=4){this.setState({error :"Turno completo, pruebe con otro o cambie de dia"})}
  }
    
  }

  //Método al que se llama cuando se elimina una clase
  private _deleteClick = async () => {

    const id = this.state.selectionId;
  
      const del = await _spService.deleteItems(id);
        if (del !== null){
          const items = await _spService.getItems();
          this.setState ( {items: items , selectionId: null, info:"Se ha borrado correctamente"});
        }
        else{
          alert("Ha habido un error");
        }
    

  }
}
