import { sp } from '@pnp/sp';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as pnp from "sp-pnp-js";
export class SPDataOperations {
  /**
	* Gets the available Choices in the Module Choice field
   *
   * @param lists The list for which the fields of type Managed Metadata need to be retrieved
   */
  public static async LOADCurrentUserTraining(lists: string, userEmail:string): Promise<any> {
    let selectedTraining: any[] = [];
    let selectedTrainingObject:any[] = [];
    let userData: any[];
    try {
      userData = await sp.web.lists.getById(lists).items.select('Created,Training/Id,Training/ModuleCalc,EmployeeID1/EMail').expand('Training,EmployeeID1').filter(`EmployeeID1/EMail eq '`+userEmail+`'`).top(500).get();
      userData[0].Training.map((training) =>{
        selectedTraining.push(training.Id);
        selectedTrainingObject.push({'Module':training.ModuleCalc,'Id':training.Id});
      })
    } catch (error) {
      console.log(error.message);
    }
    let allselectedTraining:any = {'selectedTraining':selectedTraining};
    let allselectedTrainingObject:any = {'selectedTrainingObject':selectedTrainingObject};
    let createdDate = {'Created': userData[0].Created}
    selectedTraining = {...allselectedTraining,...allselectedTrainingObject,...createdDate};
    return selectedTraining;
  }
 
    /**
   * Gets the available sub module
   *
   * @param lists The list for which the fields of type Managed Metadata need to be retrieved
   * @param module
   * @param userTrainingList
   */
  public static async LOADSubModuleData(lists: string,userEmail:any,userTrainingList:string): Promise<any> {
    let allData: any;
    let selectedTraining: any;
    let moduleData:any[] = [];
    let subModuleData:any = {};
    let trainingData:any = {};
    let trainingIDs:any = {};
    try {
      selectedTraining = await this.LOADCurrentUserTraining(userTrainingList,userEmail);
      allData = await sp.web.lists.getById(lists).items.select('Id,Title,Module,SubModule,TrainingPath').filter(`Created lt datetime'`+ selectedTraining.Created +`'`).top(500).get();
      allData.map((field) =>{
        if(moduleData.indexOf(field.Module) === -1){
          moduleData.push(field.Module);
          subModuleData[field.Module] = [];
          trainingIDs[field.Module] = [];
        }
      });
      allData.map((field) =>{
        if(subModuleData[field.Module].indexOf(field.SubModule) === -1){
          subModuleData[field.Module].push(field.SubModule);
          trainingData[field.SubModule] = [];
        }
      });
      allData.map((field) =>{
        if(trainingData[field.SubModule].indexOf(field) === -1){
          trainingData[field.SubModule].push(field);
          trainingIDs[field.Module].push(field.Id);
        }
      });
    } catch (error) {
      console.log(error.message);
    }
 
    let allSelectedTraining = {'selectedTraining':selectedTraining.selectedTraining};
    let allModuleData:any = {'module':moduleData};
    let allSubModuleData:any = {'subModule':subModuleData};
    let allTrainingData:any = {'trainingData':trainingData};
    let allTrainingIds:any = {'trainingIds':trainingIDs};
    allData = {...allModuleData,...allSubModuleData,...allTrainingData,...allSelectedTraining,...allTrainingIds}
        return allData;
  }
 
  /**
   * Check if the current user has requested permissions on a list
   *
   * @param lists The list on which user permission needs to be checked
   * @param ids The permission kind for which user needs to be authorized
   * @param itemId
   * @param pageContext
   * @param props
   * @param userAssessmentList
   */
  public static async UpdateTrainings(lists: string, trainingIds:any[], props:any, ModuleStatus?:any, userAssessmentList?: string){
    const userEmail = props.pageContext.user.email;
    const pageContext = props.pageContext;
    let SPDATA = await this.getListItemEntityType(lists);
    let userData:any = await sp.web.lists.getById(lists).items.select('Id,EmployeeID1/EMail').expand('EmployeeID1').filter(`EmployeeID1/EMail eq '`+userEmail+`'`).get();
    let itemId = userData.length>0 ? userData[0].Id : 0;
 
       let body: string = JSON.stringify({
          '__metadata': { 'type': SPDATA },
          'TrainingId': {
            'results': trainingIds
         },
         'ModuleStatus':ModuleStatus
        });
 
    props.spHttpClient.post(`${pageContext.web.absoluteUrl}/_api/web/lists/getbyid('${lists}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
      headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=verbose',
      'odata-version': '',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE'
      },
      body: body
      })
      .then((response: SPHttpClientResponse): void => {
        if(ModuleStatus !== "" && ModuleStatus !== undefined){
          this.AssignModuleAssessment(userAssessmentList,ModuleStatus,props);
        } else {
          window.location.href = window.location.pathname+"?assessmentSubmit=true";
          //window.location.reload();
        }
      }, (error: any): void => {
        console.log(error);
      });
  }
/**
   * Gets the available Choices in the Module Choice field
   * @param lists The list for which the fields of type Managed Metadata need to be retrieved
   * @param assessmentlist
   * @param totalQuestion
   * @param userEmail
   * @param userAssessentList
   */
  public static async LOADCurrentUserAssessment(lists: string, assessmentlist:string, totalQuestion:any, userEmail:string, userAssessentList: string): Promise<any> {
    let selecedModule: any;
    let assessments:any[] = [];
    let correctAnswer:any = {};
    let userAnswer:any = {};
    let assessmentAttempt: any = {};
    try {
      let userData:any = await sp.web.lists.getById(lists).items.select('ModuleStatus,EmployeeID1/EMail').expand('EmployeeID1').filter(`EmployeeID1/EMail eq '`+userEmail+`'`).get();
      selecedModule = userData.length > 0 ? userData[0].ModuleStatus : "";
 
      if(selecedModule !=""){
        assessmentAttempt = await this.GetAssessmentStatus(userAssessentList,userEmail);
 
        if(assessmentAttempt.totalAttempt === 0 || (assessmentAttempt.assessmentStatus === 'Fail' && assessmentAttempt.totalAttempt < 3)){
            assessments = await sp.web.lists.getById(assessmentlist).items.select('Id,Title,A,B,OData__x0043_,D,E,Answer').filter(`Module eq '`+encodeURIComponent(selecedModule)+`'` ).get();
            assessments.sort((a, b) => {return 0.5 - Math.random();});
 
            const selectedItems = assessments.slice(0, +totalQuestion).map(item => {
              correctAnswer[item.Id] = item.Answer;
              userAnswer[item.Id] = "";
              return item;
            });
            assessments = selectedItems;
        }
      }
    } catch (error) {
      console.log(error.message);
    }
    let assessmentModule:any = {'assessmentModule':selecedModule};
    let assessmentData:any = {'assessmentData':assessments};
    let assessmentAnswer:any = {'correctAnswer':correctAnswer};
    let assessmentQuestion:any = {'userAnswer':userAnswer};
    let assessmentTotalAttempt:any = {'totalAttempt':assessmentAttempt};
    return {...assessmentModule,...assessmentData,...assessmentAnswer,...assessmentQuestion,...assessmentTotalAttempt};
  }
 
  /**
   * Check if the current user has requested permissions on a list
   *
   * @param lists The list on which user permission needs to be checked
   * @param userEmail The permission kind for which user needs to be authorized
   */
  public static async GetAssessmentStatus(lists: string, userEmail: string){
    let assessmentAttemptData = await sp.web.lists.getById(lists).items.select('*,EmployeeID1/EMail').expand('EmployeeID1').orderBy('Modified', false).top(1).filter(`EmployeeID1/EMail eq '`+userEmail+`'`).get();
    return assessmentAttemptData.length > 0 ? {attemptId:assessmentAttemptData[0].Id,totalAttempt:assessmentAttemptData[0].Attempt,assessmentStatus:assessmentAttemptData[0].AssessmentStatus || '',assessmentAllData:assessmentAttemptData[0]} : {attemptId:0,totalAttempt:0,assessmentStatus:'',assessmentAllData:{}};
  }
 
/**
   * Check if the current user has requested permissions on a list
   *
   * @param lists The list on which user permission needs to be checked
   * @param ids The permission kind for which user needs to be authorized
   * @param itemId
   * @param pageContext
   * @param props
   */
  public static async UpdateAssessmentStatus(lists: string, module: string, status: string, totalAttemptData:any,props:any, correctQuestion:number,score:number,totalQuestion:number){
    let totalAttempt:number = totalAttemptData.totalAttempt+1;
    let SPDATA = await this.getListItemEntityType(lists);
    const body: string = JSON.stringify({
      '__metadata': { 'type': SPDATA },
      'Attempt': totalAttempt,
      'AssessmentStatus':status,
      'totalQuestion': totalQuestion,
      'passingScore': +props.passingScore,
      'correctQuestion': correctQuestion,
      'score':score.toFixed(2)
    });
 
    props.context.spHttpClient.post(`${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbyid('${lists}')/items(${totalAttemptData.attemptId})`,
      SPHttpClient.configurations.v1,
      {
      headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=verbose',
      'odata-version': '',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE'
      },
      body: body
      })
      .then(async (response: SPHttpClientResponse): Promise<void> => {
        if(totalAttempt===3 && status==="Fail"){
          let selectedTraining = await this.LOADCurrentUserTraining(props.userTrainingList,props.context.pageContext.user.email);
          let selectedTrainingObject = selectedTraining.selectedTrainingObject;
          let updatedTrainingId:any = [];
          selectedTrainingObject.map((val) => {
            if(val.Module != module){
              updatedTrainingId.push(val.Id);
            }
          });
          const updateTraining = await this.UpdateTrainings(props.userTrainingList,updatedTrainingId,props.context);
        } else {
          window.location.href = window.location.pathname+"?assessmentSubmit=true";
        }
      }, (error: any): void => {
        console.log(error);
      });
  }
 
  public static async AssignModuleAssessment(lists: string, module: string, props:any){
    let SPDATA = await this.getListItemEntityType(lists);
    let userDetails = await this.spLoggedInUserDetails(props);
    const body: string = JSON.stringify({
      '__metadata': { 'type': SPDATA },
      'Title': module,
      'EmployeeID1Id':userDetails.Id
    });
 
    props.spHttpClient.post(`${props.pageContext.web.absoluteUrl}/_api/web/lists/getbyid('${lists}')/items`,
      SPHttpClient.configurations.v1,
      {
      headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=verbose',
      'odata-version': ''
      },
      body: body
      })
      .then(async (response: SPHttpClientResponse): Promise<void> => {
        window.location.href = window.location.pathname+"?assessment=true";
      });
  }
 
    /*Get Current Logged In User*/
    public static async spLoggedInUserDetails(ctx: any): Promise<any>{
  try {
      const web = new pnp.Web(ctx.pageContext.site.absoluteUrl);
      return await web.currentUser.get();
    } catch (error) {
      console.log("Error in spLoggedInUserDetails : " + error);
    }
  }
 
   /**
   * Check if the current user has requested permissions on a list
   * @param listId The list on which user permission needs to be checked
   */
  public static async getListItemEntityType(listId: string){
    let entityType:any;
    try {
      entityType = await sp.web.lists.getById(listId).getListItemEntityTypeFullName();
    } catch(error){
      console.log('SPDataOperations.getListItemEntityType' + error);
    }
    return entityType;
  }

  public static async getListItems(listGuid: string, fields: string, expand?: string, filter?: string) {
    let allItems: any[];
    try {
      allItems = await sp.web.lists.getById(listGuid).items.select(fields).expand(expand).filter(filter).get();
    } catch(error) {
      console.log('SPDataOperations.getListItems' + error);
    }
    return allItems;
  }

  public static async updateListItem(listGuid: string, itemId: number, jsonData: any) {
    try {
      await sp.web.lists.getById(listGuid).items.getById(itemId).update(jsonData);
    } catch(error) {
      console.log('SPDataOperations.updateListItem' + error);
    }
  }

  public static async getUserID(userLoginName: string){
    let userObject: any;
    try {
      userObject = await sp.web.siteUsers.getByLoginName(userLoginName).get();
    } catch(error) {
      console.log('SPDataOperations.getUserID' + error);
    }
    return userObject;
  }

  public static async getAttachment(listGuid: string, itemId: number) {
    let allAttachment: any[];
    try {
      allAttachment = await sp.web.lists.getById(listGuid).items.getById(itemId).attachmentFiles.get();
    } catch(error) {
      console.log('SPDataOperations.getListItems' + error);
    }
    return allAttachment;
  }

  public static async addAttachment(listGuid: string, itemId: number, fileNo: number, file: any) {
    try {
      let fileName: string = file.name;
      fileName = fileNo +'.' + fileName.split('.').reverse()[0];
      await sp.web.lists.getById(listGuid).items.getById(itemId).attachmentFiles.add(fileName, file);
    } catch(error) {
      console.log('SPDataOperations.getListItems' + error);
    }
  }

  public static async deleteAttachment(listGuid: string, itemId: number, fileName: string) {
    try {
      await sp.web.lists.getById(listGuid).items.getById(itemId).attachmentFiles.getByName(fileName).delete();
    } catch(error) {
      console.log('SPDataOperations.getListItems' + error);
    }
  }

  /*Get Current Logged In User*/  
  public static async getLoggedInUserDetails(ctx: any): Promise<any>{  
    try {  
      return await sp.web.currentUser.get();   
    } catch (error) {  
      console.log("SPDataOperations.getLoggedInUserDetails " + error);  
    }      
  }

  /*Get choice column value */  
  public static async getChoicesFromChoiceColumn(listGuid: string, ColumnName: string): Promise<any>{  
    try {
      return await sp.web.lists.getById(listGuid).fields.getByInternalNameOrTitle(ColumnName).select('Choices,ID').get();
    } catch (error) {  
      console.log("SPDataOperations.getChoicesFromChoiceColumn " + error);  
    }      
  }
}