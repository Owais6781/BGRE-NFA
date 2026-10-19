var vSrAppNo1=[]; 
var vDOAMat=[];
var arraycount = 0;  
var fileUploadeCount = 0;
var vStatusSuccFailed='Failed';
var vUniqNo='';
var vAttachID=0;
$( document ).ready(function()
{   
/*alert("Please do not create a new request at this time. Workflow migration is in progress, and any new data may not be saved correctly. The process is expected to take 2 to 3 weeks.");

		if(_spPageContextInfo.userId==31971)
		{
		$('#dvAllBtns').show();
		
		}else
		{
		$('#dvAllBtns').hide();
		}*/

    debugger;
    $('title').text('Note for Approval(NFA) | New Request');
    var today= new Date();
    vUniqNo=getUniqueID();    
	var dd = today.getDate();
	var mm = today.getMonth() + 1; 
	var yyyy = today.getFullYear();
	var vCurrentDate=dd+"/"+mm +"/"+yyyy ;
    document.getElementById("spnDate").innerHTML=vCurrentDate;
    BindLoginUserProperty();
    $('#file_input').multifile();//For facilitate multi file upload    
    $.getScript("/sites/teams/NFA/SiteAssets/js/jquery.multiselect.js", function()
    { 
	   $('#ddlProjectSubEntity').multiselect
       ({
		    columns: 1,
		    placeholder: 'Select Project Sub-Entity'
	   });
    });
    debugger;    
    BindFunctions();
    BindProjects();
    $("#txtPackageAwardedCost").keypress(function()
    {
      validate(event);
    }); 
    $("#txtBudgetAmount").keypress(function()
    {
      validate(event);
    });  
    $('.content').richText();
    $('.content1').richText();
    
    
   
    
	var vChronology = "<b>1.	Complete Indent/PR received from & Date :-</b><br/>";
	vChronology    += "<b>2.	Inquiry floated on                           :-</b><br/>";
	vChronology    += "<b>3.	RFQ Ref No.            :-</b><br/>";
	vChronology    += "<b>4.	Initial Quote received on.            :-</b><br/>";
	vChronology    += "<b>5.	BOQ/Specifications Revisions by & date :- </b><br/>";
		 
	var vPaymentTerms = "<b>1.	Payment Terms :-</b><br/>";
	vPaymentTerms    += "<b>2.	Adhoc Amount :-</b><br/>";
	vPaymentTerms    += "<b>3.	Retention :- </b>(Five) 5% of the Value of the work including extra items (if any) as certified by the Owner/Owner’s representative shall be deducted in each running bill upto maximum of 5% of value. <br/>";
	vPaymentTerms    += "<b>4.	Release of Retention :- </b>Shall be released after completion of Defect Liability Period.<br/>";

	$(".content2").html(vChronology);
	document.getElementById("txtChronology").value = vChronology ;	 
	$('.content2').richText();
	
	$(".content3").html(vPaymentTerms);
	document.getElementById("txtPaymentTermsNew").value = vPaymentTerms;	 
	$('.content3').richText(); 
	
	 $('.contentIntro').richText();
    
    var option_all ="";
    $('#ddlProjectSubEntity').change(function() 
    {   
        debugger;       
        $('#hdnProjectSubEntity').val($(this).val()); 
        option_all = $(this).find("option:selected").map(function ()
        {
          return $(this).text();
        }).get().join(',');         
        document.getElementById("hdnProjectSubEntity").value =option_all;       
     });
     
});



function getUniqueID() 
{
	try
	{
		var chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXTZabcdefghiklmnopqrstuvwxyz";
		var string_length = 15;
		var randomstring = '';
		for (var i = 0; i < string_length; i++) 
		{
			var rnum = Math.floor(Math.random() * chars.length);
			randomstring += chars.substring(rnum, rnum + 1);
		}
		var res =  randomstring.toUpperCase();
	}
	catch(err)
	{
	 alert(err);
	}
    return res;
}



function validate(evt) 
{
  var theEvent = evt || window.event;
  var key = theEvent.keyCode || theEvent.which;
  key = String.fromCharCode(key);
  var regex = /^[+-]?\d*\.?\d{0,2}$/;
  if (!regex.test(key)) 
  {
      theEvent.returnValue = false;
      if (theEvent.preventDefault) theEvent.preventDefault();
  }
}

function BindFunctions() 
{    
    var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('FunctionMaster')//items?$top=500&$select=Title,ID&$filter=(IsActive eq '1')&$orderby=Title asc";
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {
            var Logg = data.d.results;            
            try 
            {
              if(data.d.results.length > 0)
              {         
                 var ddlFunctionName= document.getElementById("ddlFunctionName");
                 ddlFunctionName.options.length = 0;       
                 ddlFunctionName.options[ddlFunctionName.options.length] = new Option("-Select Function-","0");       
                 for(var i=0;i< data.d.results.length;i++)
                 {
                   ddlFunctionName.options[ddlFunctionName.options.length] = new Option(data.d.results[i].Title,data.d.results[i].ID);
                 }     
              }
	        }
            catch (e) {                               
            }
        },
        error: function () {
            console.log("error");
        }
    });   
}


function BindProjects() 
{    
    var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('Project Master')//items?$top=500&$select=Title,ProjectName,ID&$filter=(IsActive eq '1')&$orderby=ProjectName asc";
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {
            var Logg = data.d.results;            
            try 
            {
              if(data.d.results.length > 0)
              {         
                 var ddlProject= document.getElementById("ddlProject");
                 ddlProject.options.length = 0;       
                 ddlProject.options[ddlProject.options.length] = new Option("-Select Project-","0");       
                 for(var i=0;i< data.d.results.length;i++)
                 {
                   ddlProject.options[ddlProject.options.length] = new Option(data.d.results[i].Title+'-'+data.d.results[i].ProjectName,data.d.results[i].ID);
                 }     
              }
	        }
            catch (e) {                               
            }
        },
        error: function () {
            console.log("error");
        }
    });   
    
}
function BindProjectVPIncharge()
{
   
    var vProjID=document.getElementById("ddlProject").value;
    var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('Project Master')//items?$top=500&$select=ProjectHead/Id,ProjectHead/Title,VPProject/Id,VPProject/Title,NFASeries&$expand=ProjectHead,VPProject&$filter=(ID eq '"+vProjID+"')";
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {
            var Logg = data.d.results;            
            try 
            {
              if(Logg != null && Logg.length > 0)
              {         
				document.getElementById("hdnProjIncharge").value=Logg[0].ProjectHead.Title+'-'+Logg[0].ProjectHead.Id;
				document.getElementById("hdnProjVicePresident").value=Logg[0].VPProject.Title+'-'+Logg[0].VPProject.Id;
				document.getElementById("hdnProjSeries").value=Logg[0].NFASeries;
              }
              else
              {
                document.getElementById("hdnProjIncharge").value='';
				document.getElementById("hdnProjVicePresident").value='';
				document.getElementById("hdnProjSeries").value='';
              }
	        }
            catch (e) {                               
            }
        },
        error: function () {
            console.log("error");
        }
    });   

}

function BindProjectSubEntity()
{
   
    var ddlProject=document.getElementById("ddlProject").value;
    var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('SubEntity')//items?$top=500&$select=Title,EntityName,EntityID&$filter=(ProjectID eq '"+ddlProject+"')&$orderby=EntityName asc";
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {
            var Logg = data.d.results;            
            try 
            {
              if(data.d.results.length > 0)
              {         
                 var ddlProjectSubEntity= document.getElementById("ddlProjectSubEntity");
                 ddlProjectSubEntity.options.length = 0;
                 for(var i=0;i< data.d.results.length;i++)
                 {
                   ddlProjectSubEntity.options[ddlProjectSubEntity.options.length] = new Option(data.d.results[i].EntityName,data.d.results[i].EntityName);
                 }  
                 $("#ddlProjectSubEntity").multiselect("reload");   
              }
	        }
            catch (e) {                               
            }
        },
        error: function () {
            console.log("error");
        }
    });   
}

function BindPRIndent()
{
    //substringof('"+ddlProject +"',LeagelEntityNameandCode)and
    var ddlProject= $("#ddlProject option:selected").text();
    var requestUri ="https://brookfield.sharepoint.com/sites/teams/NFA/PR/_api/web/lists/getbytitle('PRProcess')//items?$top=500&$select=RequestNo,LeagelEntityNameandCode&$filter=( Status eq 'Accepted' )&$orderby=RequestNo asc";
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {
            var Logg = data.d.results;            
            try 
            {
              if(data.d.results.length > 0)
              {         
                 var ddlPrIndent= document.getElementById("ddlPrIndent");
                 ddlPrIndent.options.length = 0;       
                 ddlPrIndent.options[ddlPrIndent.options.length] = new Option("-Select PRIndent-","-Select PRIndent-");       
                 for(var i=0;i< data.d.results.length;i++)
                 {
                   ddlPrIndent.options[ddlPrIndent.options.length] = new Option(data.d.results[i].RequestNo,data.d.results[i].RequestNo);
                 }     
              }
	        }
            catch (e) {                               
            }
        },
        error: function () {
            console.log("error");
        }
    });      
    
}


function onProjectChanges()
{
    BindProjectSubEntity();
    //BindPRIndent();  
    BindProjectVPIncharge();
     
}

function onFunctionChange()
{  
	var vItemID = document.getElementById("ddlFunctionName").value;	
	BindManualApproval(vItemID);
	BindDOAMatrix(vItemID);
}

function onChangeManualDOA()
{
  
}

function onChangeDOA()
{
  var vDOA=document.getElementById("ddlDOAMatrix");
  var txtvDOA = vDOA.options[vDOA.selectedIndex].text;
  var vDOAID= vDOA.options[vDOA.selectedIndex].value;
  GetDOADetails(vDOAID);
  document.getElementById("dvDOACaption").style.display='none';
}


function GetDOADetails(vItemID) 
{   
   var vFunctionID = document.getElementById("ddlFunctionName").value;
   var requestUri ='';
   var ret = ''; 
   var requestUri ='';
   if(vFunctionID  != "0" && vFunctionID != '-Select Function-')
   { 
    requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('DOAMaster')//items?$select=*,Title,DOADetail,Level1/Name,Level1/Id,Level1/Title,Level2/Name,Level2/Id,Level2/Title,Level3/Name,Level3/Id,Level3/Title,Level4/Title,Level4/Name,Level4/Id,Level5/Name,Level5/Id,Level5/Title,Level6/Name,Level6/Id,Level6/Title,Level7/Name,Level7/Id,Level7/Title,Level8/Name,Level8/Id,Level8/Title,Level9/Name,Level9/Id,Level9/Title&$expand=Level1,Level2,Level3,Level4,Level5,Level6,Level7,Level8,Level9&$filter=(ID eq '"+vItemID+"' and FunctionID eq '"+vFunctionID+"')";
   }
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {
            var Logg = data.d.results;            
            try 
            {          
                  var strHtmlHeader="<table style='width:100%' class='blueTable' cellspacing='0'><tr><td colspan='9' style='text-align:center;'><label style='font-family:Calibri; text-align:center; font-size:small'><b>DOA Approval Matrix</b></label></td></tr>";   
                  strHtmlHeader+= "<tr>";
                  var strHtmlBody = "<tr>";   
                  var strHtmlBodyUsers="<tr>" ; 
                  ret  = data.d.results[0].Title;
                  document.getElementById("dvFinancialImpact").style.display='block';                   
                  document.getElementById("dvFinancialImpact").innerHTML="<font face='Calibri'><b>Financial Impact</b> : <b>"+data.d.results[0].ValFinancialImpact+"</b></font>";
                  if(data.d.results[0].ValFinancialImpactCon != null)
                  {                  
                    document.getElementById("hdnValFinancialImpactCon").value=data.d.results[0].ValFinancialImpactCon;
                  }
                  else
                  {
                    document.getElementById("hdnValFinancialImpactCon").value='';
                  }
                  if(data.d.results[0].ValFinImpactMin != null)
                  {                  
                    document.getElementById("hdnValFinancialImpactMin").value=data.d.results[0].ValFinImpactMin;
                  }
                  else
                  {
                    document.getElementById("hdnValFinancialImpactMin").value='';
                  }

                  
                 if(data.d.results[0].Level1.Id !=null && data.d.results[0].Level1.Id !="")
                 {
                   if(data.d.results[0].L1Name == "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL1").value=vProjIncharge;
                   }
                   else if(data.d.results[0].L1Name == "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL1").value=vProjVP;
                   }
                   else
                   {
                     document.getElementById("hdnDOAL1").value=data.d.results[0].Level1.Id;
                   }
                 }
                 else
                 { 
                   if(data.d.results[0].L1Name == "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL1").value=vProjIncharge;
                   }
                   else if(data.d.results[0].L1Name == "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL1").value=vProjVP;
                   }
                   else
                   {
                     document.getElementById("hdnDOAL1").value="";
                   }
                 }
                 if(data.d.results[0].Level2.Id !=null && data.d.results[0].Level2.Id !="")
                 {
                   if(data.d.results[0].L2Name == "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL2").value=vProjIncharge;
                   }
                   else if(data.d.results[0].L2Name == "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL2").value=vProjVP;
                   }
                   else
                   {
                   document.getElementById("hdnDOAL2").value=data.d.results[0].Level2.Id;
                   }
                 }
                 else
                 { 
                   if(data.d.results[0].L2Name == "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL2").value=vProjIncharge;
                   }
                   else if(data.d.results[0].L2Name == "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL2").value=vProjVP;
                   }
                   else
                   {
                     document.getElementById("hdnDOAL2").value="";
                   }
                 }

                 if(data.d.results[0].Level3.Id !=null && data.d.results[0].Level3.Id !="")
                 {
                   if(data.d.results[0].Level3Name== "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL3").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level3Name== "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL3").value=vProjVP;
                   }
                   else
                   {
                   document.getElementById("hdnDOAL3").value=data.d.results[0].Level3.Id;
                   }
                 }
                 else
                 {
                   if(data.d.results[0].Level3Name== "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL3").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level3Name== "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL3").value=vProjVP;
                   }
                   else
                   {
                     document.getElementById("hdnDOAL3").value="";
                   }
                 }

                 if(data.d.results[0].Level4.Id !=null && data.d.results[0].Level4.Id !="")
                 {
                   if(data.d.results[0].Level4Name== "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL4").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level4Name== "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL4").value=vProjVP;
                   }
                   else
                   {
                   document.getElementById("hdnDOAL4").value=data.d.results[0].Level4.Id;
                   }
                 }
                 else
                 { 
                   if(data.d.results[0].Level4Name== "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL4").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level4Name== "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL4").value=vProjVP;
                   }
                   else
                   {
                    document.getElementById("hdnDOAL4").value="";
                   }
                 }

                 if(data.d.results[0].Level5.Id !=null && data.d.results[0].Level5.Id !="")
                 {
                   if(data.d.results[0].Level5Name== "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL5").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level5Name== "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL5").value=vProjVP;
                   }
                   else
                   {
                    document.getElementById("hdnDOAL5").value=data.d.results[0].Level5.Id;
                   }
                 }
                 else
                 {
                   if(data.d.results[0].Level5Name== "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL5").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level5Name== "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL5").value=vProjVP;
                   }
                   else
                   {
                      document.getElementById("hdnDOAL5").value="";
                   }
                 }

                 if(data.d.results[0].Level6.Id !=null && data.d.results[0].Level6.Id !="")
                 {
                   if(data.d.results[0].Level6Name== "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL6").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level6Name== "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL6").value=vProjVP;
                   }
                   else
                   {
	                  document.getElementById("hdnDOAL6").value=data.d.results[0].Level6.Id;
                   }
                 }
                 else
                 {
	               if(data.d.results[0].Level6Name== "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL6").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level6Name== "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL6").value=vProjVP;
                   }
                   else
                   {
                     document.getElementById("hdnDOAL6").value="";
                   }
                 }

                 if(data.d.results[0].Level7.Id !=null && data.d.results[0].Level7.Id !="")
                 {
                   if(data.d.results[0].Level7Name== "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL7").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level7Name== "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL7").value=vProjVP;
                   }
                   else
                   {
                   document.getElementById("hdnDOAL7").value=data.d.results[0].Level7.Id;
                   }
                 }
                 else
                 {
                   if(data.d.results[0].Level7Name== "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL7").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level7Name== "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL7").value=vProjVP;
                   }
                   else
                   {
                      document.getElementById("hdnDOAL7").value="";
                   }
                 }

                 if(data.d.results[0].Level8.Id !=null && data.d.results[0].Level8.Id !="")
                 {
                   if(data.d.results[0].Level8Name == "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL8").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level8Name == "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL8").value=vProjVP;
                   }
                   else
                   {
                     document.getElementById("hdnDOAL8").value=data.d.results[0].Level8.Id;
                   }
                 }
                 else
                 {
                   if(data.d.results[0].Level8Name == "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL8").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level8Name == "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL8").value=vProjVP;
                   }
                   else
                   {
                    document.getElementById("hdnDOAL8").value="";
                   }
                 }

                 if(data.d.results[0].Level9.Id !=null && data.d.results[0].Level9.Id !="")
                 {
                   if(data.d.results[0].Level9Name== "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL9").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level9Name== "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL9").value=vProjVP;
                   }
                   else
                   {
                     document.getElementById("hdnDOAL9").value=data.d.results[0].Level9.Id;
                   }
                 }
                 
                 else
                 {
                   if(data.d.results[0].Level9Name == "Project Incharge")
                   {
                     var vProjIncharge = document.getElementById("hdnProjIncharge").value.split('-')[1];
                     document.getElementById("hdnDOAL9").value=vProjIncharge;
                   }
                   else if(data.d.results[0].Level9Name == "VP Projects")
                   {
                     var vProjVP = document.getElementById("hdnProjVicePresident").value.split('-')[1];
                     document.getElementById("hdnDOAL9").value=vProjVP;
                   }
                   else
                   {
                     document.getElementById("hdnDOAL9").value="";
                   }
                 }
                 document.getElementById("hdnDOAL10").value="";
                 document.getElementById("hdnDOAL11").value="";
                 document.getElementById("hdnDOAL12").value="";
                  
                  if(data.d.results[0].L1Name != null)
                  {
                    strHtmlHeader += "<td ><font face='Calibri' size='2'><b>L-1</b></font></td>";
                    strHtmlBody += "<td ><font face='Calibri' size='2'>"+data.d.results[0].L1Name+"</font></td>";
                    if(data.d.results[0].Level1.Title != null)
                    {
	                     if(data.d.results[0].L1Name== "Project Incharge")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else if(data.d.results[0].L1Name== "VP Projects")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else
	                     {
	                      strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+data.d.results[0].Level1.Title+"</b></font></td>";
	                     }
                    }
                    else
		            {
						 if(data.d.results[0].L1Name== "Project Incharge")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else if(data.d.results[0].L1Name== "VP Projects")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
						else
						{						
						   strHtmlBodyUsers += "<td><font face='Calibri' size='2'></font></td>";
						}
		            }
                  }
                                   
                  if(data.d.results[0].L2Name != null)
                  {
                    strHtmlHeader+= "<td ><font face='Calibri' size='2'><b>L-2</b></font></td>";
                    strHtmlBody += "<td ><font face='Calibri' size='2'>"+data.d.results[0].L2Name +"</font></td>";                  
                    if(data.d.results[0].Level2.Title != null)
                    {
                         if(data.d.results[0].L2Name == "Project Incharge")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else if(data.d.results[0].L2Name == "VP Projects")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else
	                     {
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+data.d.results[0].Level2.Title+"</b></font></td>";
	                     }
                    }
                    else
		            {
						if(data.d.results[0].L2Name == "Project Incharge")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else if(data.d.results[0].L2Name == "VP Projects")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
						else
						{						
						    strHtmlBodyUsers += "<td><font face='Calibri' size='2'></font></td>";
						}
		            }
                  }
                 

                  
                  if(data.d.results[0].Level3Name != null)
                  {
                    strHtmlHeader+= "<td ><font face='Calibri' size='2'><b>L-3</b></font></td>";
                    strHtmlBody += "<td ><font face='Calibri' size='2'>"+data.d.results[0].Level3Name+"</font></td>";
                    if(data.d.results[0].Level3.Title != null)
                    {
                         if(data.d.results[0].Level3Name == "Project Incharge")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else if(data.d.results[0].Level3Name == "VP Projects")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else
	                     {
	                      strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+data.d.results[0].Level3.Title+"</b></font></td>";
	                     }
                    }
                    else
		            {
						if(data.d.results[0].Level3Name == "Project Incharge")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else if(data.d.results[0].Level3Name == "VP Projects")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
						else
						{						
						    strHtmlBodyUsers += "<td><font face='Calibri' size='2'></font></td>";
						}
		            }
                  }
                 

                  
                  if(data.d.results[0].Level4Name != null)
                  {
	                    strHtmlHeader+= "<td ><font face='Calibri' size='2'><b>L-4</b></font></td>";
	                    strHtmlBody += "<td ><font face='Calibri' size='2'>"+data.d.results[0].Level4Name +"</font></td>";
	                    if(data.d.results[0].Level4.Title != null)
	                    {
			                 if(data.d.results[0].Level4Name == "Project Incharge")
		                     {
		                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
		                     }
		                     else if(data.d.results[0].Level4Name == "VP Projects")
		                     {
		                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
		                     }
		                     else
		                     {
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+data.d.results[0].Level4.Title+"</b></font></td>";
		                     }
	                    }
	                    else
		                {
						     if(data.d.results[0].Level4Name == "Project Incharge")
		                     {
		                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
		                     }
		                     else if(data.d.results[0].Level4Name == "VP Projects")
		                     {
		                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
		                     }
							 else
							 {							
								strHtmlBodyUsers += "<td><font face='Calibri' size='2'></font></td>";
							 }
		                }
                  }
                  

                  
                  if(data.d.results[0].Level5Name  != null)
                  {
	                   strHtmlHeader+= "<td><font face='Calibri' size='2'><b>L-5</b></font></td>";
	                   strHtmlBody += "<td ><font face='Calibri' size='2'>"+data.d.results[0].Level5Name+"</font></td>";
	                   if(data.d.results[0].Level5.Title != null)
	                   {
		                     if(data.d.results[0].Level5Name == "Project Incharge")
		                     {
		                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
		                     }
		                     else if(data.d.results[0].Level5Name == "VP Projects")
		                     {
		                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
		                     }
		                     else
		                     {
	                           strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+data.d.results[0].Level5.Title+"</b></font></td>";
	                         }
	                   }
	                   else
	                   {
							 if(data.d.results[0].Level5Name == "Project Incharge")
		                     {
		                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
		                     }
		                     else if(data.d.results[0].Level5Name == "VP Projects")
		                     {
		                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
		                     }
							 else
							 {							
								strHtmlBodyUsers += "<td><font face='Calibri' size='2'></font></td>";
							 }
	                   }	                   
                  }
                  

                  
                  if(data.d.results[0].Level6Name != null)
                  {
	                   strHtmlHeader+= "<td><font face='Calibri' size='2'><b>L-6</b></font></td>";
	                   strHtmlBody += "<td ><font face='Calibri' size='2'>"+data.d.results[0].Level6Name+"</font></td>";
	                   if(data.d.results[0].Level6.Title != null)
	                   {
	                        if(data.d.results[0].Level6Name == "Project Incharge")
		                     {
		                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
		                     }
		                     else if(data.d.results[0].Level6Name == "VP Projects")
		                     {
		                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
		                     }
		                     else
		                     {
	                           strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+data.d.results[0].Level6.Title+"</b></font></td>";
	                         }
	                   }
	                   else
	                   {
							 if(data.d.results[0].Level6Name == "Project Incharge")
		                     {
		                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
		                     }
		                     else if(data.d.results[0].Level6Name == "VP Projects")
		                     {
		                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
		                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
		                     }
							 else
							 {							
								strHtmlBodyUsers += "<td><font face='Calibri' size='2'></font></td>";
							 }
	                   }
                  }
                 

                  if(data.d.results[0].Level7Name != null)
                  {
                    strHtmlHeader+= "<td ><font face='Calibri' size='2'><b>L-7</b></font></td>";
                    strHtmlBody += "<td ><font face='Calibri' size='2'>"+data.d.results[0].Level7Name +"</font></td>";
                    if(data.d.results[0].Level7.Title != null)
	                {
	                     if(data.d.results[0].Level7Name == "Project Incharge")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else if(data.d.results[0].Level7Name == "VP Projects")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else
	                     {
	                        strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+data.d.results[0].Level7.Title+"</b></font></td>";
	                     }
	                }
	                else
                    {
						if(data.d.results[0].Level7Name == "Project Incharge")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else if(data.d.results[0].Level7Name== "VP Projects")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
						else
						{						
						 	strHtmlBodyUsers += "<td><font face='Calibri' size='2'></font></td>";
						}
                    }
                  }
                  

                  
                  if(data.d.results[0].Level8Name  != null)
                  {
                    strHtmlHeader+= "<td ><font face='Calibri' size='2'><b>L-8</b></font></td>";
                    strHtmlBody += "<td ><font face='Calibri' size='2'>"+data.d.results[0].Level8Name+"</font></td>";
                    if(data.d.results[0].Level8.Title != null)
	                {
	                     if(data.d.results[0].Level8Name == "Project Incharge")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else if(data.d.results[0].Level8Name == "VP Projects")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else
	                     {
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+data.d.results[0].Level8.Title+"</b></font></td>";
	                     }
	                }
	                else
                    {
						if(data.d.results[0].Level8Name == "Project Incharge")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else if(data.d.results[0].Level8Name == "VP Projects")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
						else
						{	
							strHtmlBodyUsers += "<td><font face='Calibri' size='2'></font></td>";
						}
                    }
                  }
                  

                  
                  if(data.d.results[0].Level9Name  != null)
                  {
                    strHtmlHeader+= "<td ><font face='Calibri' size='2'><b>L-9</b></font></td>";
                    strHtmlBody += "<td><font face='Calibri' size='2'>"+data.d.results[0].Level9Name +"</font></td>";
                    if(data.d.results[0].Level9.Title != null)
	                {    
	                     if(data.d.results[0].Level9Name == "Project Incharge")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else if(data.d.results[0].Level9Name == "VP Projects")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else
	                     {
	                        strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+data.d.results[0].Level9.Title+"</b></font></td>";
	                     }
	                }
	                else
                    {
						if(data.d.results[0].Level9Name == "Project Incharge")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjIncharge").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
	                     else if(data.d.results[0].Level9Name == "VP Projects")
	                     {
	                       var hdnHOD=document.getElementById("hdnProjVicePresident").value.split('-')[0];
	                       strHtmlBodyUsers += "<td><font face='Calibri' size='2'><b>"+hdnHOD+"</b></font></td>";
	                     }
						else
						{						
							strHtmlBodyUsers += "<td><font face='Calibri' size='2'></font></td>";
						}
                    }
                  }
                 

                  strHtmlHeader += "</tr>"; 
                  strHtmlBody += "</tr>"; 
                  strHtmlBodyUsers += "</tr>";
                  document.getElementById("dvDOAMatrixData").style.display='block';
                  document.getElementById("dvDOAMatrixData").innerHTML=strHtmlHeader+strHtmlBody+strHtmlBodyUsers;
	        }
            catch (e) {
                ret = '';                
            }
        },
        error: function () {
            console.log("error");
        }
    });
    return ret;
}


function BindDOAMatrix(varItemID) 
{    
    var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('DOAMaster')//items?$top=5000&$select=ID,Title,DOADetail,ValFinancialImpact&$filter=(FunctionID eq '"+varItemID+"')&$orderby=Title asc";
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {
            var Logg = data.d.results;            
            try 
            {
              if(data.d.results.length > 0)
              {         
                 var ddlDOAMatrix= document.getElementById("ddlDOAMatrix");
                 ddlDOAMatrix.options.length = 0;       
                 ddlDOAMatrix.options[ddlDOAMatrix.options.length] = new Option("-Select DOA Clause-","0");       
                 for(var i=0;i< data.d.results.length;i++)
                 {
                   ddlDOAMatrix.options[ddlDOAMatrix.options.length] = new Option(data.d.results[i].ID+' - '+data.d.results[i].Title +'('+data.d.results[i].ValFinancialImpact+')',data.d.results[i].ID);                    
                 }
              }
              else
              {
                 var ddlDOAMatrix= document.getElementById("ddlDOAMatrix");
                 ddlDOAMatrix.options.length = 0;
              }
	        }
            catch (e) {                               
            }
        },
        error: function () {
            console.log("error");
        }
    });  
    
}


function BindLoginUserProperty() 
{    
    debugger;
    var vCurrentUser="i:0%23.f|membership|"+_spPageContextInfo.userEmail;
    var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"+vCurrentUser+"'";
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {                      
            try 
            {
              debugger;
              if(data.d)
              { 
                 var Logg = data.d.UserProfileProperties.results;  
				 for (i = 0; i < Logg.length; i++)
				 {  
					  if (Logg[i].Key == "Department")
					  {
					    document.getElementById("hdnDepartment").value=Logg[i].Value;
					  }					  
					  if (Logg[i].Key == "Office")
					  {
					    document.getElementById("hdnLocation").value=Logg[i].Value;
					  }
				 }
			  }
	        }
            catch (e) {                               
            }
        },
        error: function () {
            console.log("error");
        }
    });  
    
}

function BindManualApproval(varItemID) 
{    
    debugger;
    var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('ApproverMaster')//items?$top=500&$select=ApproverName/Name,ApproverName/Title,ApproverName/Id&$expand=ApproverName&$filter=(FunctionID eq '"+varItemID+"' and IsActive eq '1')";
    var vManArr=[];
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {
            var Logg = data.d.results;            
            try 
            {
              
              var ddlApprover1= document.getElementById("ddlApprover1");
                 ddlApprover1.options.length = 0;       
                 ddlApprover1.options[ddlApprover1.options.length] = new Option("-Select-","0");   
                 
                 var ddlApprover2= document.getElementById("ddlApprover2");
                 ddlApprover2.options.length = 0;       
                 ddlApprover2.options[ddlApprover2.options.length] = new Option("-Select-","0"); 
                 
                 var ddlApprover3= document.getElementById("ddlApprover3");
                 ddlApprover3.options.length = 0;       
                 ddlApprover3.options[ddlApprover3.options.length] = new Option("-Select-","0");  
                 
				 var ddlApprover4= document.getElementById("ddlApprover4");
				 ddlApprover4.options.length = 0;       
				 ddlApprover4.options[ddlApprover4.options.length] = new Option("-Select-","0");  
				 
				 var ddlApprover5= document.getElementById("ddlApprover5");
				 ddlApprover5.options.length = 0;       
				 ddlApprover5.options[ddlApprover5.options.length] = new Option("-Select-","0"); 
				 
				 var ddlApprover6= document.getElementById("ddlApprover6");
				 ddlApprover6.options.length = 0;       
				 ddlApprover6.options[ddlApprover6.options.length] = new Option("-Select-","0");
				 
				 var ddlApprover7= document.getElementById("ddlApprover7");
				 ddlApprover7.options.length = 0;       
				 ddlApprover7.options[ddlApprover7.options.length] = new Option("-Select-","0");
				 
				 var ddlApprover8= document.getElementById("ddlApprover8");
				 ddlApprover8.options.length = 0;       
				 ddlApprover8.options[ddlApprover8.options.length] = new Option("-Select-","0");
				 
				 var ddlApprover9= document.getElementById("ddlApprover9");
				 ddlApprover9.options.length = 0;       
				 ddlApprover9.options[ddlApprover9.options.length] = new Option("-Select-","0");
				 
				 var ddlApprover10= document.getElementById("ddlApprover10");
				 ddlApprover10.options.length = 0;       
				 ddlApprover10.options[ddlApprover10.options.length] = new Option("-Select-","0");
				 
				 var ddlApprover11= document.getElementById("ddlApprover11");
				 ddlApprover11.options.length = 0;       
				 ddlApprover11.options[ddlApprover11.options.length] = new Option("-Select-","0");
				 
				 var ddlApprover12= document.getElementById("ddlApprover12");
				 ddlApprover12.options.length = 0;       
				 ddlApprover12.options[ddlApprover12.options.length] = new Option("-Select-","0");
	              if(data.d.results.length > 0)
	              { 
	                 for(var i=0;i< data.d.results.length;i++)
	                 {
	                   var VApprovalName=data.d.results[i].ApproverName.Title;
	                   vManArr.push(VApprovalName+'|'+data.d.results[i].ApproverName.Id);
	                 }  
	                 vManArr= array_unique(vManArr.sort());
                     for(var j=0; j < vManArr.length; j++)
                     {                      
	                   ddlApprover1.options[ddlApprover1.options.length]   = new Option(vManArr[j].split('|')[0],vManArr[j].split('|')[1]);
	                   ddlApprover2.options[ddlApprover2.options.length]   = new Option(vManArr[j].split('|')[0],vManArr[j].split('|')[1]);
	                   ddlApprover3.options[ddlApprover3.options.length]   = new Option(vManArr[j].split('|')[0],vManArr[j].split('|')[1]);
	                   ddlApprover4.options[ddlApprover4.options.length]   = new Option(vManArr[j].split('|')[0],vManArr[j].split('|')[1]);
	                   ddlApprover5.options[ddlApprover5.options.length]   = new Option(vManArr[j].split('|')[0],vManArr[j].split('|')[1]);
	                   ddlApprover6.options[ddlApprover6.options.length]   = new Option(vManArr[j].split('|')[0],vManArr[j].split('|')[1]);
	                   ddlApprover7.options[ddlApprover7.options.length]   = new Option(vManArr[j].split('|')[0],vManArr[j].split('|')[1]);
	                   ddlApprover8.options[ddlApprover8.options.length]   = new Option(vManArr[j].split('|')[0],vManArr[j].split('|')[1]);
	                   ddlApprover9.options[ddlApprover9.options.length]   = new Option(vManArr[j].split('|')[0],vManArr[j].split('|')[1]);
	                   ddlApprover10.options[ddlApprover10.options.length] = new Option(vManArr[j].split('|')[0],vManArr[j].split('|')[1]);
	                   ddlApprover11.options[ddlApprover11.options.length] = new Option(vManArr[j].split('|')[0],vManArr[j].split('|')[1]);
	                   ddlApprover12.options[ddlApprover12.options.length] = new Option(vManArr[j].split('|')[0],vManArr[j].split('|')[1]);
                     }   
	              }
	        }
            catch (e) {                               
            }
        },
        error: function () {
            console.log("error");
        }
    });  
    
}


function array_unique(array) {
    var unique = [];
    for ( var i = 0 ; i < array.length ; ++i ) {
        if ( unique.indexOf(array[i]) == -1 )
            unique.push(array[i]);
    }
    return unique;
}

function OnManualApproval()
{  
       
    var vSrAppNo=[];
    var ddlApprover1= document.getElementById("ddlApprover1").value;
    var ddlApprover2= document.getElementById("ddlApprover2").value;
    var ddlApprover3= document.getElementById("ddlApprover3").value;
    var ddlApprover4= document.getElementById("ddlApprover4").value;
    var ddlApprover5= document.getElementById("ddlApprover5").value;
    var ddlApprover6= document.getElementById("ddlApprover6").value;
    var ddlApprover7= document.getElementById("ddlApprover7").value;
    var ddlApprover8= document.getElementById("ddlApprover8").value;
    var ddlApprover9= document.getElementById("ddlApprover9").value;
    var ddlApprover10= document.getElementById("ddlApprover10").value;
    var ddlApprover11= document.getElementById("ddlApprover11").value;
    var ddlApprover12= document.getElementById("ddlApprover12").value;
    
    if(ddlApprover1 != "0")
    {
      document.getElementById("hdnDOAL1").value=ddlApprover1;
      vSrAppNo.push(1);
    }
    else
    {
      document.getElementById("hdnDOAL1").value="";
    }
    
    if(ddlApprover2 != "0")
    {
      document.getElementById("hdnDOAL2").value=ddlApprover2;
      vSrAppNo.push(2);
    }
    else
    {
      document.getElementById("hdnDOAL2").value="";
    }

    
    if(ddlApprover3 != "0")
    {
      document.getElementById("hdnDOAL3").value=ddlApprover3;
      vSrAppNo.push(3);
    } else
    {
      document.getElementById("hdnDOAL3").value="";
    }

    
    if(ddlApprover4 != "0")
    {
      document.getElementById("hdnDOAL4").value=ddlApprover4;
      vSrAppNo.push(4);
    } else
    {
      document.getElementById("hdnDOAL4").value="";
    }

    
    if(ddlApprover5 != "0")
    {
      document.getElementById("hdnDOAL5").value=ddlApprover5;
      vSrAppNo.push(5);
    } else
    {
      document.getElementById("hdnDOAL5").value="";
    }

    
    if(ddlApprover6 != "0")
    {
      document.getElementById("hdnDOAL6").value=ddlApprover6;
      vSrAppNo.push(6);
    } else
    {
      document.getElementById("hdnDOAL6").value="";
    }

    
    if(ddlApprover7 != "0")
    {
      document.getElementById("hdnDOAL7").value=ddlApprover7;
      vSrAppNo.push(7);
    } else
    {
      document.getElementById("hdnDOAL7").value="";
    }

    
    if(ddlApprover8 != "0")
    {
      document.getElementById("hdnDOAL8").value=ddlApprover8;
      vSrAppNo.push(8);
    } else
    {
      document.getElementById("hdnDOAL8").value="";
    }

    
    if(ddlApprover9 != "0")
    {
      document.getElementById("hdnDOAL9").value=ddlApprover9;
      vSrAppNo.push(9);
    } else
    {
      document.getElementById("hdnDOAL9").value="";
    }


    if(ddlApprover10 != "0")
    {
      document.getElementById("hdnDOAL10").value=ddlApprover10;
      vSrAppNo.push(10);
    } else
    {
      document.getElementById("hdnDOAL10").value="";
    }

    
    if(ddlApprover11 != "0")
    {
      document.getElementById("hdnDOAL11").value=ddlApprover11;
      vSrAppNo.push(11);
    } else
    {
      document.getElementById("hdnDOAL11").value="";
    }

    
    if(ddlApprover12 != "0")
    {
      document.getElementById("hdnDOAL12").value=ddlApprover12;
      vSrAppNo.push(12);
    } else
    {
      document.getElementById("hdnDOAL12").value="";
    }
    vSrAppNo1=vSrAppNo;
}
function GetUserName(varID) 
{
    
    var ret = null;    
    var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/getuserbyid('"+varID+"')";
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {
            var Logg = data.d;            
            try 
            {
                
	            ret=Logg.Title; 
	        }
            catch (e) {
                ret = null;                
            }
        },
        error: function () {
            console.log("error");
        }
    });
    return ret;
}

function GetCurrentDataToday()
{
    var today= new Date();    
	var dd = today.getDate();
	var mm = today.getMonth() + 1; 
	var yyyy = today.getFullYear();
	var vDueDate=dd+"/"+mm +"/"+yyyy ;
	
	var hours = today.getHours();
	var minutes = today.getMinutes();
	var ampm = hours >= 12 ? 'PM' : 'AM';
	hours = hours % 12;
	hours = hours ? hours : 12; // the hour '0' should be '12'
	minutes = minutes < 10 ? '0'+minutes : minutes;
	var strTime = vDueDate+' '+hours + ':' + minutes + ' ' + ampm;
    return strTime ;
}

function SubmitApprovalNotes()
{
   var vOK='OK';var vCount=1;
   var vWorkflowStartUp=document.getElementById("ddlWorkflowType").value;
   if(ValidateUploadNotes())
   {
       if(vWorkflowStartUp == 'MANUAL'  || vWorkflowStartUp == 'Manual')
       {
           for(var i=0;i<vSrAppNo1.length;i++)
           {
             if(vCount != vSrAppNo1[i])
             {
               vOK='No';
               break;
             }
             vCount = parseInt(vCount)+1;
           }
       } 
       if(vOK=='OK')
       {
           onn();
		   document.getElementById("dvAllBtns").style.display='none'; 			  
		   formSave();	
	   }  
	   else
	   {
	     var msg = "<font face='Calibri'><b>Please select manual approver in a sequence (e.g : Level-1, Level-2, Level-3).</b></font>"; 
	     warningalert1(msg ,'');
	   } 
   }
}
function GetUserName(varID) 
{
    
    var ret = null;    
    var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/getuserbyid('"+varID+"')";
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {
            var Logg = data.d;            
            try 
            {
                
	            ret=Logg.Title; 
	        }
            catch (e) {
                ret = null;                
            }
        },
        error: function () {
            console.log("error");
        }
    });
    return ret;
}


function GetCounter(varFunctionID) 
{    
    var vFunAbb='';
    debugger;
    var vRetVal='';
    var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('CounterMaster')//items?$top=500&$select=ID,Title,Counter,FunctionAbbr,FunctionID";
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {
             debugger;

            var Logg = data.d.results;            
            try 
            {            
				var dtt = new Date();
				var mm = dtt.getMonth(); 
				var yyyy = dtt.getFullYear();
				var month = new Array();
				month[0] = "Jan";
				month[1] = "Feb";
				month[2] = "Mar";
				month[3] = "Apr";
				month[4] = "May";
				month[5] = "Jun";
				month[6] = "Jul";
				month[7] = "Aug";
				month[8] = "Sep";
				month[9] = "Oct";
				month[10] = "Nov";
				month[11] = "Dec";
				var mName = month[mm];
	            if(Logg.length > 0)
	            {
		             var hdnProjSeries = document.getElementById('hdnProjSeries').value;
		             for(var i=0; i < Logg.length;i++)
		             {
		                 if(Logg[i].FunctionID == varFunctionID)
		                 {
		                   vFunAbb=Logg[i].FunctionAbbr;
		                 }
		                 if(Logg[i].FunctionID == '0')
		                 {				            
				             document.getElementById("hdnCounter").value=Logg[i].Counter;
				             document.getElementById("hdnCounterItemID").value=Logg[i].ID;
			             }
		             }
		             if(vFunAbb != '')
		             {
		              	vRetVal=hdnProjSeries+'/'+vFunAbb+'/'+yyyy+'/'+mName+'/'+document.getElementById("hdnCounter").value;
		              	 document.getElementById("hdnDigitalRequestNo").value=vRetVal;
		             }
		             else
		             {
		              	vRetVal=hdnProjSeries+'/'+yyyy+'/'+mName+'/'+document.getElementById("hdnCounter").value;
		              	 document.getElementById("hdnDigitalRequestNo").value=vRetVal;
		             }
	            }
	        }
            catch (e) {                               
            }
        },
        error: function (data) {
             debugger;
            console.log(data.responseText);
        }
    });  
    return vRetVal;
}



 function UpdateCounter() 
 {
		debugger;
		var itemId = document.getElementById("hdnCounterItemID").value; 
		var hdnCounter= document.getElementById("hdnCounter").value;  
		hdnCounter=parseInt(hdnCounter)+1;
        $.ajax({
        
            url: _spPageContextInfo.webAbsoluteUrl +"/_api/web/lists/getByTitle('CounterMaster')/getItemByStringId('" +itemId + "')",
            type: "POST",
            contentType: "application/json;odata=verbose",
            async: false,
            data: JSON.stringify(
            {
                '__metadata': {
                    'type': 'SP.Data.CounterMasterListItem'
                },
                'Counter': hdnCounter.toString()              
            }),
            headers: {
                "accept": "application/json;odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                "IF-MATCH": "*",
                "X-Http-Method": "PATCH"
            },
            success: function (data) 
            {
                AddDashBoard(document.getElementById("hdnDigitalRequestNo").value,document.getElementById("hdnApprvlNotesItemID").value);               
            },
            error: function (e) 
            {
            }
        });
}






function ValidateUploadNotes()
{
    debugger;
    var msg = "";
    var focusId="";
    var rtvalue = true; 
    
    var ddlFunctionName  = $('#ddlFunctionName').val();
    var ddlProject       = $('#ddlProject').val();
    var ddlWorkflowType  = $('#ddlWorkflowType').val();
    var ddlDOAMatrix     = $('#ddlDOAMatrix').val();
    var txtSubjectMatter = $('#txtSubjectMatter').val();
    var txtIntrPreBack   = $('#txtIntrPreBack').val();
    var txtCommercialProposal = $('#txtCommercialProposal').val();
    var txtPaymentTermsNew    = $('#txtPaymentTermsNew').val();  
    var txtRetention          = $('#txtRetention').val();    
    var txtRefundRetention    = $('#txtRefundRetention').val();
    var txtLiquidatedDamage   = $('#txtLiquidatedDamage').val();
    var txtContractPerformGurantee = $('#txtContractPerformGurantee').val();    
    var txtDefectLibility          = $('#txtDefectLibility').val();
    var txtWorkCompletionPeriod= $('#txtWorkCompletionPeriod').val();
    var ddlBudgetedNonBudgeted = $('#ddlBudgetedNonBudgeted').val(); 
    var txtJustification       = $('#txtJustification').val(); 
    var txtPackageAwardedCost  = $('#txtPackageAwardedCost').val();   
    var txtContractorName      = $('#txtContractorName').val();
    var txtBudgetAmount        = $('#txtBudgetAmount').val();

    var vCount=0; var vFileSpclChar=0;
    if (ddlFunctionName== '' || ddlFunctionName=="0" || ddlFunctionName=="-Select Function-") 
    {  
        msg += "<font face='Calibri'><b>Please select Originator Department.</b></font>" + "<br/>";
        if (focusId== "") 
        {
         focusId=  $('#ddlFunctionName');
        }		
    }     

   
    if (ddlProject== '' || ddlProject == "0" || ddlProject == "-Select Project-") 
    {   
        msg += "<font face='Calibri'><b>Please select Project.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#ddlProject');	
        } 	
    }
    
    if (txtSubjectMatter== '') 
    {   
        msg += "<font face='Calibri'><b>Please enter Reference / Subject.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#txtSubjectMatter');	
        } 	
    }
    
    if (txtIntrPreBack== '') 
    {   
        msg += "<font face='Calibri'><b>Please enter Introduction / Preamble / Background.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#txtIntrPreBack');	
        } 	
    }
    
    if (txtCommercialProposal== '') 
    {   
        msg += "<font face='Calibri'><b>Please enter Commercial Proposal/s.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#txtCommercialProposal');	
        } 	
    }
    
    if (txtPaymentTermsNew == '') 
    {   
        msg += "<font face='Calibri'><b>Please enter Payment Terms.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#txtPaymentTermsNew');	
        } 	
    }
    
    
    
    if (txtLiquidatedDamage== '') 
    {   
        msg += "<font face='Calibri'><b>Please enter Liquidated Damage.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#txtLiquidatedDamage');	
        } 	
    }
    
    if (txtContractPerformGurantee== '') 
    {   
        msg += "<font face='Calibri'><b>Please enter Contract Performance Guarantee.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#txtContractPerformGurantee');	
        } 	
    }
    
    if (txtDefectLibility== '') 
    {   
        msg += "<font face='Calibri'><b>Please enter Defect Libility Period.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#txtDefectLibility');	
        } 	
    }
    
    if (txtWorkCompletionPeriod== '') 
    {   
        msg += "<font face='Calibri'><b>Please enter Work Completion Period.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#txtWorkCompletionPeriod');	
        } 	
    }
    
    if (ddlBudgetedNonBudgeted == 'Budgeted') 
    {   
        if(parseFloat(txtBudgetAmount) > 0)
        { 	       	
        }
        else
        {
            msg += "<font face='Calibri'><b>Please enter Budget Amount.</b></font>" + "<br/>";        
	        if (focusId== "") 
	        {
	          focusId=  $('#txtBudgetAmount');	
	        } 
        }
    }
    
    if(txtContractorName == "")
    {
        msg += "<font face='Calibri'><b>Please enter Contractor Name.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#txtContractorName');	
        }
    }
    
    if (txtJustification== '') 
    {   
        msg += "<font face='Calibri'><b>Please enter Recommendation with justification.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#txtJustification');	
        } 	
    }
    if (ddlWorkflowType != '' && ddlWorkflowType != "0" && ddlWorkflowType != "-Select Workflow Type-") 
    {  
       
       if(ddlWorkflowType == "MANUAL")
       {
          var vCountManualAppr=0;
          var ddlApprover1=$('#ddlApprover1').val();
          var ddlApprover2=$('#ddlApprover2').val();
          var ddlApprover3=$('#ddlApprover3').val();
          var ddlApprover4=$('#ddlApprover4').val();
          var ddlApprover5=$('#ddlApprover5').val();
          var ddlApprover6=$('#ddlApprover6').val();
          var ddlApprover7=$('#ddlApprover7').val();
          var ddlApprover8=$('#ddlApprover8').val();
          var ddlApprover9=$('#ddlApprover9').val();
          var ddlApprover10=$('#ddlApprover10').val();
          var ddlApprover11=$('#ddlApprover11').val();
          var ddlApprover12=$('#ddlApprover12').val();
          
          if(ddlApprover1 != null && ddlApprover1 !='0')
          {
            vCountManualAppr = parseInt(vCountManualAppr)+1;
          }
          if(ddlApprover2 != null && ddlApprover2 !='0')
          {
            vCountManualAppr = parseInt(vCountManualAppr)+1;
          }
          if(ddlApprover3 != null && ddlApprover3 !='0')
          {
            vCountManualAppr = parseInt(vCountManualAppr)+1;
          }
          if(ddlApprover4 != null && ddlApprover4 !='0')
          {
            vCountManualAppr = parseInt(vCountManualAppr)+1;
          }
          if(ddlApprover5 != null && ddlApprover5 !='0')
          {
            vCountManualAppr = parseInt(vCountManualAppr)+1;
          }
          if(ddlApprover6 != null && ddlApprover6 !='0')
          {
            vCountManualAppr = parseInt(vCountManualAppr)+1;
          }
          if(ddlApprover7 != null && ddlApprover7 !='0')
          {
            vCountManualAppr = parseInt(vCountManualAppr)+1;
          }
          if(ddlApprover8 != null && ddlApprover8 !='0')
          {
            vCountManualAppr = parseInt(vCountManualAppr)+1;
          }
          if(ddlApprover9 != null && ddlApprover9 !='0')
          {
            vCountManualAppr = parseInt(vCountManualAppr)+1;
          }
          if(ddlApprover10 != null && ddlApprover10 !='0')
          {
            vCountManualAppr = parseInt(vCountManualAppr)+1;
          }
          if(ddlApprover11 != null && ddlApprover11 !='0')
          {
            vCountManualAppr = parseInt(vCountManualAppr)+1;
          }
          if(ddlApprover12 != null && ddlApprover12 !='0')
          {
            vCountManualAppr = parseInt(vCountManualAppr)+1;
          }
          
          if(parseInt(vCountManualAppr)== 0)
          {
			msg += "<font face='Calibri'><b>Please select Approval level-1.</b></font><br/>";        
			if (focusId== "") 
			{
			  focusId=  $('#ddlApprover1');	
			} 	
          }
       }
       else if(ddlWorkflowType == "DOA")
       {
         if(ddlDOAMatrix == null || ddlDOAMatrix   == undefined || ddlDOAMatrix == '' || ddlDOAMatrix == '-Select DOA Clause-' || ddlDOAMatrix == '0')
         {
            document.getElementById("ddlWorkflowType").value="MANUAL";
            document.getElementById("dvDOACaption").style.display='none';
            document.getElementById("dvManualApproval").style.display='block';
            msg += "<font face='Calibri'><b>Please select DOA Matrix.</b></font><br/>";        
			if (focusId== "") 
			{
			  focusId=  $('#ddlDOAMatrix');	
			} 
         }
         else
         {
           var hdnDOAL1= document.getElementById("hdnDOAL1").value;
           var hdnDOAL2= document.getElementById("hdnDOAL2").value;
           var hdnDOAL3= document.getElementById("hdnDOAL3").value;
           if(hdnDOAL1 == '')
           {
	            document.getElementById("ddlWorkflowType").value="MANUAL";
	            document.getElementById("dvDOACaption").style.display='none';
	            document.getElementById("dvManualApproval").style.display='block';
	            msg += "<font face='Calibri'><b>Selected DOA having no approver.Please select another DOA.</b></font><br/>";        
				if (focusId== "") 
				{
				  focusId=  $('#ddlDOAMatrix');	
				}
           }   
         }
       }
    }
    
     
     
     if ($('#dvAttachment').html() == "" ) 
     {   
        msg += "<font face='Calibri'><b>Please attach Supporting document.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#file_input');	
        } 	
     }
     
    
    
    if (msg != "")
    {
      warningalert1(msg ,focusId);
      rtvalue = false;
	  return false;
    }   
    return rtvalue;
}

function cancelbtn()
{
   debugger;
   window.location="https://brookfield.sharepoint.com/sites/teams/NFA/SitePages/Home.aspx";
}
   
   
   
function AddDashBoard(vRequestNo,vWFID)
{  
    debugger;
    var vLevel="1";
    var vStatus='Pending';        
	   
    var vFunction=document.getElementById("ddlFunctionName");
	var txtvFunction = vFunction.options[vFunction.selectedIndex].text;
	var vIDFunction= vFunction.options[vFunction.selectedIndex].value;	
    var ddlWorkflowType=document.getElementById("ddlWorkflowType").value;
    var txtAmount=document.getElementById("txtPackageAwardedCost").value;
    var txtPaymentTerms=document.getElementById("txtPaymentTermsNew").value;
    var ddlProject=document.getElementById("ddlProject");   
	var txtProject = ddlProject.options[ddlProject.selectedIndex].text;	   
    var hdnLocation=document.getElementById("hdnLocation").value;
    var hdnDepartment=document.getElementById("hdnDepartment").value;

    var ddlDOAMatrix=document.getElementById("ddlDOAMatrix");
	var txtDOAClause ="NA";
	if(ddlDOAMatrix.value != null && ddlDOAMatrix.value != '' && ddlDOAMatrix.value != "0" && ddlDOAMatrix.value != '-Select DOA Clause-')
	{
		txtDOAClause = ddlDOAMatrix.options[ddlDOAMatrix.selectedIndex].text;	   
	}

    var hdnDOAL1="0"; var hdnCurrApprName='';
    if(document.getElementById("hdnDOAL1").value !="" && ddlWorkflowType == 'MANUAL')
    {    
       hdnDOAL1=document.getElementById("hdnDOAL1").value;
       var ddlApprover1=document.getElementById("ddlApprover1");
	   var hdnDOAL1Text= ddlApprover1.options[ddlApprover1.selectedIndex].text;	  
    }
    else if(document.getElementById("hdnDOAL1").value !="" && ddlWorkflowType == 'DOA')
    {
      hdnDOAL1=document.getElementById("hdnDOAL1").value;
    }
   
    var vCurrentDate=GetCurrentDataToday();
    
	var someDate = new Date();	        	               
	var dd = someDate.getDate();
	var mm = someDate.getMonth() + 1;
	var yyyy = someDate.getFullYear();
	if (dd < 10) 
	{
	  dd = '0' + dd;
	}
	if (mm < 10) 
	{
	  mm = '0' + mm;
	}	
	var someFormattedDate = mm + '/'+ dd + '/'+ yyyy;
	
	var vDueDate = new Date(someFormattedDate);
	vDueDate .setDate(vDueDate .getDate() + 1);	
	 
	var dd1 = vDueDate.getDate();
	var mm1 = vDueDate.getMonth() + 1;
	var yyyy1 = vDueDate.getFullYear();
	if (dd1 < 10) 
	{
	  dd1 = '0' + dd1;
	}
	if (mm1 < 10) 
	{
	  mm1 = '0' + mm1;
	}
    var vMainDueDate=mm1 + '/'+ dd1 + '/'+ yyyy1;
    var url = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('DashboardList')/items";
    var data = JSON.stringify(  
      {  
        "__metadata":   
        {  
            "type": "SP.Data.DashboardListListItem"  
        },  
        "Title": vRequestNo, 
        "Function": txtvFunction ,
        "Project": txtProject,
        "Amount" : txtAmount,   
        "DOA":txtDOAClause,           
        "WorkflowType": ddlWorkflowType,
        "CurrentApproverId": hdnDOAL1,       
        "CurrentApprovalID": hdnDOAL1,  
        "Status": vStatus,
        "WFLevel": vLevel,
        "PendingFrom": vCurrentDate,
        "PendingFromDate": someFormattedDate,		            
		"InitiatedBy":_spPageContextInfo.userDisplayName,
		"InitiatedOn":vCurrentDate,  
        "FunctionID": vIDFunction,     
        "WFID": vWFID,        
		"Location": hdnLocation,
		"Department": hdnDepartment,
		"DueDate": vMainDueDate        
    });  
    $.ajax  
    ({  
        url: url,  
        type: "POST",  
        data: data, 
        async: false, 
        headers: {  
            "Accept": "application/json;odata=verbose",  
            "Content-Type": "application/json;odata=verbose",  
            "X-RequestDigest": $("#__REQUESTDIGEST").val()
              
        },  
        success: function(data)   
      	{  
      	    debugger;   
      	    CreateLogsAfetrApprovedPending('Pending','1',vRequestNo,vWFID);   
      	    InsertDataIntoReminderList(vWFID,vRequestNo);
      	    off();	   
		   	alert("Note for Approval has been Submitted for approval. Your request no. is : "+document.getElementById("hdnDigitalRequestNo").value);
	        window.location="/sites/teams/NFA/SitePages/Home.aspx";		               
        },  
        error: function(data)  
        {  
            debugger;
            alert("Failed to create list item.");  
        }  
    });
}

function InsertDataIntoReminderList(vWFID,vNFANo)
{
      var url = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('WFReminderAlert')/items";
      var data = JSON.stringify(  
      {  
        "__metadata":   
        {  
            "type": "SP.Data.WFReminderAlertListItem"  
        },  
        "Title": vNFANo, 
        "WFID": vWFID,
        "Status":'New'
      });  
    $.ajax  
    ({  
        url: url,  
        type: "POST",  
        data: data, 
        async: false, 
        headers: {  
            "Accept": "application/json;odata=verbose",  
            "Content-Type": "application/json;odata=verbose",  
            "X-RequestDigest": $("#__REQUESTDIGEST").val()
              
        },  
        success: function(data)   
      	{ 	               
        },  
        error: function(data)  
        {            
        }  
    });
}

function CreateLogsAfetrApprovedPending(vWFStatus,vLevel,vRequestNo,vWFID)
{  
    debugger;	     
	  
    var vFunction=document.getElementById("ddlFunctionName");
	var txtvFunction = vFunction.options[vFunction.selectedIndex].text;
	var vIDFunction= vFunction.options[vFunction.selectedIndex].value;
	
    var ddlWorkflowType=document.getElementById("ddlWorkflowType").value;
    var txtAmount=document.getElementById("txtPackageAwardedCost").value;
    var ddlProject=document.getElementById("ddlProject");   
	var txtProject = ddlProject.options[ddlProject.selectedIndex].text;	   
    var hdnLocation=document.getElementById("hdnLocation").value;
    var hdnDepartment=document.getElementById("hdnDepartment").value;
    
    var hdnDOAL1="0"; var hdnCurrApprName='';
    if(document.getElementById("hdnDOAL1").value !="" && ddlWorkflowType == 'MANUAL')
    {    
       hdnDOAL1=document.getElementById("hdnDOAL1").value;
       var ddlApprover1=document.getElementById("ddlApprover1");
	   var hdnDOAL1Text= ddlApprover1.options[ddlApprover1.selectedIndex].text;	  
    }
    else if(document.getElementById("hdnDOAL1").value !="" && ddlWorkflowType == 'DOA')
    {
      hdnDOAL1=document.getElementById("hdnDOAL1").value;
    }
    var ddlDOAMatrix=document.getElementById("ddlDOAMatrix");
	var txtDOAClause ="NA";
	if(ddlDOAMatrix.value != null && ddlDOAMatrix.value != '')
	{
		txtDOAClause = ddlDOAMatrix.options[ddlDOAMatrix.selectedIndex].text;	   
	}
    
    var url = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('WFLogList')/items";
    var data = JSON.stringify(  
      {  
        "__metadata":   
        {  
            "type": "SP.Data.WFLogListListItem"  
        },  
        "Title": vRequestNo, 
        "WFID": vWFID,
        "FunctionID": vIDFunction,
        "FunctionName" : txtvFunction,
        "DOAClause":txtDOAClause,       
        "Project": txtProject,
        "Amount": txtAmount,   
        "Status": vWFStatus,       
        "PPLPendingAtId": hdnDOAL1,		          
		"WFLevel":vLevel,
		"Location":hdnLocation,  
        "Department": hdnDepartment
    });  
    $.ajax  
    ({  
        url: url,  
        type: "POST",  
        data: data, 
        async: false, 
        headers: {  
            "Accept": "application/json;odata=verbose",  
            "Content-Type": "application/json;odata=verbose",  
            "X-RequestDigest": $("#__REQUESTDIGEST").val()
              
        },  
        success: function(data)   
      	{   
      	     	 
        },  
        error: function(data)  
        {  
            debugger;
            alert("Failed to create list item.");  
        }  
    });
}

 
    
function replaceAll(str, find, replace) 
{
    return str.replace(new RegExp(find, 'g'), replace);
}

    
    
 var options;
 function openModalDialog(vURL,vTitleData) 
 {        
    options = SP.UI.$create_DialogOptions();
    options.width = 1200;
    options.height = 450;        
    options.title='View DOA Approval Matrix';       
    options.url = vURL;
    options.dialogReturnValueCallback = Function.createDelegate(null, CloseCallback);   
    SP.SOD.execute('sp.ui.dialog.js', 'SP.UI.ModalDialog.showModalDialog', options);
 }
function CloseCallback(result, target) {
}


function includesMethod(container, value) {
	var returnValue = false;
	for(var i=0;i<value.length;i++)
	{
		var pos = container.indexOf(value[i]);
		if (pos >= 0) {
			returnValue = true;
			break;
		}
	}
	return returnValue;
}



  
  
    
function formSave() 
{  
    var listItem = "";  
    var fileArray = [];  
    
	var vLevel="1";
	var vStatus='Pending';
	
	
	var vFunction=document.getElementById("ddlFunctionName");
	var txtvFunction = vFunction.options[vFunction.selectedIndex].text;
	var vIDFunction= vFunction.options[vFunction.selectedIndex].value;	
	var ddlProject=document.getElementById("ddlProject");
	var txtvProject = ddlProject.options[ddlProject.selectedIndex].text;
	var vIDProject= ddlProject.options[ddlProject.selectedIndex].value;	
	var ddlProjectSubEntity= document.getElementById("hdnProjectSubEntity").value;
	var  txtPaymentTerms= document.getElementById("txtPaymentTermsNew").value;
	var txtCommercialProposal=document.getElementById("txtCommercialProposal").value;
	var txtChronology = document.getElementById("txtChronology").value;
	var txtSubjectMatter= document.getElementById("txtSubjectMatter").value;
	var txtIntrPreBack= document.getElementById("txtIntrPreBack").value;   
    var txtLiquidatedDamage= document.getElementById("txtLiquidatedDamage").value;    
    var txtContractPerformGurantee= document.getElementById("txtContractPerformGurantee").value;
    var txtDefectLibility= document.getElementById("txtDefectLibility").value;
    var txtWorkCompletionPeriod= document.getElementById("txtWorkCompletionPeriod").value;    
    var txtBudgetedNonBudgeted= document.getElementById("ddlBudgetedNonBudgeted").value;
    var txtPackageAwardedCost=document.getElementById("txtPackageAwardedCost").value;
    var txtContractorName=document.getElementById("txtContractorName").value;
    var txtJustification= document.getElementById("txtJustification").value;      
    var txtAnnexure= document.getElementById("txtAnnexure").value;    
    var txtAdd1           = document.getElementById("txtAdd1").value;
    var txtAdd2           = document.getElementById("txtAdd2").value;
    var txtAdd3           = document.getElementById("txtAdd3").value;
    var txtAdd4           = document.getElementById("txtAdd4").value;
    var ddlWorkflowType   = document.getElementById("ddlWorkflowType").value;
    var ddlNatureOfApprvr = document.getElementById("ddlNatureOfApprvr").value;    
	var hdnLocation     = document.getElementById("hdnLocation").value;
	var hdnDepartment   = document.getElementById("hdnDepartment").value;	
	var txtBudgetAmount = document.getElementById("txtBudgetAmount").value;
	var vTitle          = GetCounter(vIDFunction); 	
	var hdnDOAL1="0" ;var hdnDOAL2="0" ;var hdnDOAL3="0" ;var hdnDOAL4="0" ;var hdnDOAL5="0" ;var hdnDOAL6="0" ;
	var hdnDOAL7="0" ;var hdnDOAL8="0" ;var hdnDOAL9="0" ;var hdnDOAL10="0" ;var hdnDOAL11="0" ;var hdnDOAL12="0" ;
	if(document.getElementById("hdnDOAL1").value !="")
	{    
		hdnDOAL1=document.getElementById("hdnDOAL1").value;
	}
	if(document.getElementById("hdnDOAL2").value !="")
	{    
		hdnDOAL2=document.getElementById("hdnDOAL2").value;
	}
	
	if(document.getElementById("hdnDOAL3").value !="")
	{    
		hdnDOAL3=document.getElementById("hdnDOAL3").value;
	}
	
	if(document.getElementById("hdnDOAL4").value !="")
	{    
		hdnDOAL4=document.getElementById("hdnDOAL4").value;
	}
	
	if(document.getElementById("hdnDOAL5").value !="")
	{    
		hdnDOAL5=document.getElementById("hdnDOAL5").value;
	}
	
	if(document.getElementById("hdnDOAL6").value !="")
	{    
	  hdnDOAL6=document.getElementById("hdnDOAL6").value;
	}
	if(document.getElementById("hdnDOAL7").value !="")
	{    
		hdnDOAL7=document.getElementById("hdnDOAL7").value;
	}
	if(document.getElementById("hdnDOAL8").value !="")
	{    
		hdnDOAL8=document.getElementById("hdnDOAL8").value;
	}
	if(document.getElementById("hdnDOAL9").value !="")
	{    
		hdnDOAL9=document.getElementById("hdnDOAL9").value;
	}
	if(document.getElementById("hdnDOAL10").value !="")
	{    
		hdnDOAL10=document.getElementById("hdnDOAL10").value;
	}
	if(document.getElementById("hdnDOAL11").value !="")
	{    
		hdnDOAL11=document.getElementById("hdnDOAL11").value;
	}
	if(document.getElementById("hdnDOAL12").value !="")
	{    
		hdnDOAL12=document.getElementById("hdnDOAL12").value;
	}
	var ddlDOAMatrix=document.getElementById("ddlDOAMatrix");
	var txtBudgetRemarks=document.getElementById("txtBudgetRemarks").value;

	var txtDOAClause ="NA";
	if(ddlDOAMatrix.value != null && ddlDOAMatrix.value != '' && ddlDOAMatrix.value != "0" && ddlDOAMatrix.value != '-Select DOA Clause-')
	{
		txtDOAClause = ddlDOAMatrix.options[ddlDOAMatrix.selectedIndex].text;	   
	}
	
	var vCurrentDateOn=GetCurrentDataToday();
	var VworkflowHistory="Pending at "+GetUserName(hdnDOAL1)+" at level-1 on "+vCurrentDateOn+"<br/>Initiated by "+_spPageContextInfo.userDisplayName+" on "+vCurrentDateOn;
    debugger;
    arraycount = fileArray.length;  
    listItem = {  
		            __metadata: { "type": "SP.Data.DigitalApprovalNotesListItem" },		            
					"Title": vTitle,
					"FunctionName": txtvFunction ,
					"FunctionID": vIDFunction,	
					"DOAClause":txtDOAClause,
					"ProjectSubEntity": ddlProjectSubEntity,				
					"WorkflowType": ddlWorkflowType, 					
					"Project": txtvProject,   
					"ProjectSubEntity": ddlProjectSubEntity,
					"SubjectMatter": txtSubjectMatter,										
					"BriefDesc": txtIntrPreBack,
					"CommercialProposal": txtCommercialProposal,
					"PaymentTerms": txtPaymentTerms,				
					"LiquidatedDamage": txtLiquidatedDamage,
					"ContractPerformance": txtContractPerformGurantee,
					"DefectLibilityPeriod": txtDefectLibility, 
					"WorkCompletionPeriod": txtWorkCompletionPeriod,					
					"BudgetedNonBudgted": txtBudgetedNonBudgeted,
					"RecommendationJust": txtJustification,	
					"ContractorName": txtContractorName,
					"ApprovalAmount":txtPackageAwardedCost,
					"AnnexureEnclosure": txtAnnexure,
					"AdditionalText1": txtAdd1,
					"AdditionalText2": txtAdd2,
                    "AdditionalText3": txtAdd3,
                    "AdditionalText4": txtAdd4,				
					"NatureofApproval": ddlNatureOfApprvr,            
					"Status": vStatus,					
					"WFLevel": vLevel,     
					"ApprovalLevel1Id": hdnDOAL1,        
					"ApprovalLevel2Id": hdnDOAL2,        
					"ApprovalLevel3Id": hdnDOAL3,       
					"ApprovalLevel4Id": hdnDOAL4,       
					"ApprovalLevel5Id": hdnDOAL5,
					"ApprovalLevel6Id": hdnDOAL6,       
					"ApprovalLevel7Id": hdnDOAL7,       
					"ApprovalLevel8Id": hdnDOAL8,        
					"ApprovalLevel9Id": hdnDOAL9,        
					"ApprovalLevel10Id": hdnDOAL10,        
					"ApprovalLevel11Id": hdnDOAL11,        
					"ApprovalLevel12Id": hdnDOAL12,
					"Budget": txtBudgetAmount,
					"Dept": hdnDepartment,
					"Location": hdnLocation,
					"WorkflowApprovalHistory": VworkflowHistory,
					"WorkflowHistory": VworkflowHistory,
					"Chronology":txtChronology,
					"BudgetRemarks":txtBudgetRemarks,
					"UniqID":vUniqNo			            
                };  
    createNewItem("DigitalApprovalNotes", fileArray , listItem);  
}  
	function createNewItem(listname, filearray, listItem)
	{  
	    debugger;
	    var dfd = $.Deferred();  
	    var initializePermsCall = PostAjax(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/GetByTitle('" + listname + "')/items", listItem);  
	    $.when(initializePermsCall).then(function (permData)
	    {  
	       debugger;
	        var id = permData.d.Id;  
	        document.getElementById("hdnApprvlNotesItemID").value=permData.d.Id;
	        UpdateCounter();
	    });
	  
	}  
    function PostAjax(siteurl, listItem) 
    {  
        debugger;
        return $.ajax({  
            url: siteurl,  
            type: "POST",  
            async: false,
            contentType: "application/json;odata=verbose",  
            data: JSON.stringify(listItem),  
            headers: {  
                "Accept": "application/json;odata=verbose", 
                "X-RequestDigest": $("#__REQUESTDIGEST").val()    
            },  
	        success: function(data)   
	      	{   
	      	     	 
	        },  
	        error: function(data)  
	        {  
	            debugger;
	            alert(data.responseText);  
	            off();
	        }  

        });  
    }  
    function loopFileUpload(listName, id, filearray, fileCount) 
    {  
        debugger;
        var dfd = $.Deferred();
        UploadAttachment(listName, id, filearray[fileCount].Attachment); 
        return dfd.promise();  
    }  
    
    
	function UploadAttachment(listname, ID, file) 
	{  
		var getFileBuffer1= function (file)
		{  		
			var deferred = $.Deferred();  
			var reader = new FileReader(); 		
			reader.onload = function (e)
			{  
			    deferred.resolve(e.target.result);  
			}  
			reader.onerror = function (e) 
			{  
			    deferred.reject(e.target.error);  
			}  
			reader.readAsArrayBuffer(file);  
			return deferred.promise();  
	    };  
	
	    getFileBuffer1(file).then(function (buffer) 
	    {  
	        $.ajax({  
	            url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('" + listname + "')/items(" + ID + ")/AttachmentFiles/add(FileName='" + file.name + "')",  
	            method: 'POST',  
	            async: false,  
	            data: buffer,  
	            processData: false,  
	            headers: {  
	                "Accept": "application/json; odata=verbose",  
	                "content-type": "application/json; odata=verbose",	              
	                "X-RequestDigest": document.getElementById("__REQUESTDIGEST").value  
	
	            }, success: onAttachmentSucess  ,error:OnAttachmentError
	        }); 
	    });  
	    function onAttachmentSucess(data) 
	    {  
	        fileUploadeCount++;	  
	        vStatusSuccFailed='Success';        
	        if (arraycount == fileUploadeCount) 
	        {  
	            console.log(data + ' uploaded successfully'); 
	            vStatusSuccFailed='Success';
	            
	        }  
	    } 
	    function OnAttachmentError()
	    {
	      console.log('Error during file upload');
	      var vItemID=document.getElementById("hdnApprvlNotesItemID").value;
	      RemoveListItem(vItemID);	     	
	      off();
	      alert('Data could not be saved. Please refresh the browser and try again.');
	      window.location='/sites/teams/NFA/SitePages/NewNFA.aspx';
	    } 
	}  
    
 function RemoveListItem(vItemID) 
 {
    var vretval='';
    debugger;
    $.ajax({  
        url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('DigitalApprovalNotes')/items(" + vItemID+ ")",  
        type: "POST",  
        async: false,
        cache: false,
        contentType: "application/json;odata=verbose",  
        headers: {  
            "Accept": "application/json;odata=verbose",             
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),  
            "IF-MATCH": "*",  
            "X-HTTP-Method": "DELETE",  
        },  
        success: function(data)
        {        
           vretval='Done';
        },  
        error: function(data) {
          vretval='Error';          
        }  
    });
    return vretval;
} 


function onChangeManualDOA()
{  
  var vItemID = document.getElementById("ddlFunctionName").value;
  var ddlWorkflowType=document.getElementById("ddlWorkflowType").value;
  if(ddlWorkflowType == "MANUAL")
  {    
    document.getElementById("dvDOACaption").style.display='none';
    document.getElementById("dvManualApproval").style.display='block';
    BindManualApproval(vItemID);
  }  
  else
  {
    document.getElementById("dvDOACaption").style.display='block';
	var vDOAMatrix=document.getElementById("ddlDOAMatrix").value;
	if(vDOAMatrix!= null && vDOAMatrix!= undefined && vDOAMatrix != '')
	{
	    var vDOA=document.getElementById("ddlDOAMatrix");
		var txtvDOA = vDOA.options[vDOA.selectedIndex].text;
		var vDOAID= vDOA.options[vDOA.selectedIndex].value;
		GetDOADetails(txtvDOA);
	    document.getElementById("dvManualApproval").style.display='none';
    }

  }
}


function SaveAsDraft()
{
  if(ValidateDraftNFA())
  {
  	formSaveDraft();
  }
}


function formSaveDraft() 
{  
    var listItem = "";  
    var fileArray = [];  
    
	var vLevel="0";
	var vStatus='Draft';
	
	
	var vFunction=document.getElementById("ddlFunctionName");
	var txtvFunction = vFunction.options[vFunction.selectedIndex].text;
	var vIDFunction= vFunction.options[vFunction.selectedIndex].value;	
	var ddlProject=document.getElementById("ddlProject");
	var txtvProject = ddlProject.options[ddlProject.selectedIndex].text;
	var vIDProject= ddlProject.options[ddlProject.selectedIndex].value;	
	var ddlProjectSubEntity= document.getElementById("hdnProjectSubEntity").value;
	var  txtPaymentTerms= document.getElementById("txtPaymentTermsNew").value;
	var txtCommercialProposal=document.getElementById("txtCommercialProposal").value;
	var txtChronology = document.getElementById("txtChronology").value;
	var txtSubjectMatter= document.getElementById("txtSubjectMatter").value;
	var txtIntrPreBack= document.getElementById("txtIntrPreBack").value; 
    var txtLiquidatedDamage= document.getElementById("txtLiquidatedDamage").value;    
    var txtContractPerformGurantee= document.getElementById("txtContractPerformGurantee").value;
    var txtDefectLibility= document.getElementById("txtDefectLibility").value;
    var txtWorkCompletionPeriod= document.getElementById("txtWorkCompletionPeriod").value;    
    var txtBudgetedNonBudgeted= document.getElementById("ddlBudgetedNonBudgeted").value;
    var txtPackageAwardedCost=document.getElementById("txtPackageAwardedCost").value;
    var txtContractorName=document.getElementById("txtContractorName").value;
    var txtJustification= document.getElementById("txtJustification").value;        
    var txtAnnexure= document.getElementById("txtAnnexure").value;    
    var txtAdd1           = document.getElementById("txtAdd1").value;
    var txtAdd2           = document.getElementById("txtAdd2").value;
    var txtAdd3           = document.getElementById("txtAdd3").value;
    var txtAdd4           = document.getElementById("txtAdd4").value;
    var ddlWorkflowType   = document.getElementById("ddlWorkflowType").value;
    var ddlNatureOfApprvr = document.getElementById("ddlNatureOfApprvr").value;    
	var hdnLocation     = document.getElementById("hdnLocation").value;
	var hdnDepartment   = document.getElementById("hdnDepartment").value;	
	var txtBudgetAmount = document.getElementById("txtBudgetAmount").value;
	var vTitle          = GetCounterDraft(); 	
	var hdnDOAL1="0" ;var hdnDOAL2="0" ;var hdnDOAL3="0" ;var hdnDOAL4="0" ;var hdnDOAL5="0" ;var hdnDOAL6="0" ;
	var hdnDOAL7="0" ;var hdnDOAL8="0" ;var hdnDOAL9="0" ;var hdnDOAL10="0" ;var hdnDOAL11="0" ;var hdnDOAL12="0" ;
	if(document.getElementById("hdnDOAL1").value !="")
	{    
		hdnDOAL1=document.getElementById("hdnDOAL1").value;
	}
	if(document.getElementById("hdnDOAL2").value !="")
	{    
		hdnDOAL2=document.getElementById("hdnDOAL2").value;
	}
	
	if(document.getElementById("hdnDOAL3").value !="")
	{    
		hdnDOAL3=document.getElementById("hdnDOAL3").value;
	}
	
	if(document.getElementById("hdnDOAL4").value !="")
	{    
		hdnDOAL4=document.getElementById("hdnDOAL4").value;
	}
	
	if(document.getElementById("hdnDOAL5").value !="")
	{    
		hdnDOAL5=document.getElementById("hdnDOAL5").value;
	}
	
	if(document.getElementById("hdnDOAL6").value !="")
	{    
	  hdnDOAL6=document.getElementById("hdnDOAL6").value;
	}
	if(document.getElementById("hdnDOAL7").value !="")
	{    
		hdnDOAL7=document.getElementById("hdnDOAL7").value;
	}
	if(document.getElementById("hdnDOAL8").value !="")
	{    
		hdnDOAL8=document.getElementById("hdnDOAL8").value;
	}
	if(document.getElementById("hdnDOAL9").value !="")
	{    
		hdnDOAL9=document.getElementById("hdnDOAL9").value;
	}
	if(document.getElementById("hdnDOAL10").value !="")
	{    
		hdnDOAL10=document.getElementById("hdnDOAL10").value;
	}
	if(document.getElementById("hdnDOAL11").value !="")
	{    
		hdnDOAL11=document.getElementById("hdnDOAL11").value;
	}
	if(document.getElementById("hdnDOAL12").value !="")
	{    
		hdnDOAL12=document.getElementById("hdnDOAL12").value;
	}
	var ddlDOAMatrix=document.getElementById("ddlDOAMatrix");
	var txtBudgetRemarks=document.getElementById("txtBudgetRemarks").value;

	var txtDOAClause ="NA";
	if(ddlDOAMatrix.value != null && ddlDOAMatrix.value != '' && ddlDOAMatrix.value != "0" && ddlDOAMatrix.value != '-Select DOA Clause-')
	{
		txtDOAClause = ddlDOAMatrix.options[ddlDOAMatrix.selectedIndex].text;	   
	}
	
	var vCurrentDateOn=GetCurrentDataToday();	
    $("#attachFilesHolder input:file").each(function () 
    {  
        if ($(this)[0].files[0]) {  
            fileArray.push({ "Attachment": $(this)[0].files[0] });  
        }  
    });  
    arraycount = fileArray.length;  
    listItem = {  
		            __metadata: { "type": "SP.Data.DigitalApprovalNotesListItem" },		            
					"Title": vTitle,
					"FunctionName": txtvFunction ,
					"FunctionID": vIDFunction,	
					"DOAClause":txtDOAClause,
					"ProjectSubEntity": ddlProjectSubEntity,				
					"WorkflowType": ddlWorkflowType, 					
					"Project": txtvProject,   
					"ProjectSubEntity": ddlProjectSubEntity,
					"SubjectMatter": txtSubjectMatter,									
					"BriefDesc": txtIntrPreBack,
					"CommercialProposal": txtCommercialProposal,
					"PaymentTerms": txtPaymentTerms,					
					"ContractPerformance": txtContractPerformGurantee,
					"LiquidatedDamage": txtLiquidatedDamage,
					"DefectLibilityPeriod": txtDefectLibility, 
					"WorkCompletionPeriod": txtWorkCompletionPeriod,					
					"BudgetedNonBudgted": txtBudgetedNonBudgeted,
					"RecommendationJust": txtJustification,					
					"ContractorName": txtContractorName,
					"ApprovalAmount":txtPackageAwardedCost,
					"AnnexureEnclosure": txtAnnexure,
					"AdditionalText1": txtAdd1,
					"AdditionalText2": txtAdd2,
                    "AdditionalText3": txtAdd3,
                    "AdditionalText4": txtAdd4,				
					"NatureofApproval": ddlNatureOfApprvr,            
					"Status": vStatus,					
					"WFLevel": vLevel,     
					"ApprovalLevel1Id": hdnDOAL1,        
					"ApprovalLevel2Id": hdnDOAL2,        
					"ApprovalLevel3Id": hdnDOAL3,       
					"ApprovalLevel4Id": hdnDOAL4,       
					"ApprovalLevel5Id": hdnDOAL5,
					"ApprovalLevel6Id": hdnDOAL6,       
					"ApprovalLevel7Id": hdnDOAL7,       
					"ApprovalLevel8Id": hdnDOAL8,        
					"ApprovalLevel9Id": hdnDOAL9,        
					"ApprovalLevel10Id": hdnDOAL10,        
					"ApprovalLevel11Id": hdnDOAL11,        
					"ApprovalLevel12Id": hdnDOAL12,
					"Budget": txtBudgetAmount,
					"Dept": hdnDepartment,
					"Location": hdnLocation,
					"Chronology":txtChronology,
					"BudgetRemarks":txtBudgetRemarks,
					"UniqID":vUniqNo
										            
                };  
    createNewItemDraft("DigitalApprovalNotes", fileArray , listItem);  
}  
function createNewItemDraft(listname, filearray, listItem)
{  
    debugger;
    var dfd = $.Deferred();  
    var initializePermsCall = PostAjax(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/GetByTitle('" + listname + "')/items", listItem);  
    $.when(initializePermsCall).then(function (permData)
    {  
        debugger;
        var id = permData.d.Id;  
        if(id > 0)
        {
	        document.getElementById("hdnApprvlNotesItemID").value=permData.d.Id;
	        if (filearray.length != 0)
	        {  
	            for (i = 0; i < filearray.length; i++)
	            {
	                loopFileUploadDraft(listname, id, filearray, i).then
	                (  
	                    function(){vStatusSuccFailed='Success';},  
	                    function (sender, args) { console.log("Error uploading");
	                    vStatusSuccFailed='Success'; 
	                    dfd.reject(sender, args); 
	                });	                             
	            }  
	        } 
	        else
	        {
	          UpdateCounterDraft();
	        } 
        }
    });  
}  

function loopFileUploadDraft(listName, id, filearray, fileCount) 
    {  
        debugger;
        var dfd = $.Deferred();
        UploadAttachmentDraft(listName, id, filearray[fileCount].Attachment); 
        return dfd.promise();  
    }  
    
    
	function UploadAttachmentDraft(listname, ID, file) 
	{  
		var getFileBuffer2= function (file)
		{  		
			var deferred = $.Deferred();  
			var reader = new FileReader(); 		
			reader.onload = function (e)
			{  
			    deferred.resolve(e.target.result);  
			}  
			reader.onerror = function (e) 
			{  
			    deferred.reject(e.target.error);  
			}  
			reader.readAsArrayBuffer(file);  
			return deferred.promise();  
	    };  
	
	    getFileBuffer2(file).then(function (buffer) 
	    {  
	        $.ajax({  
	            url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('" + listname + "')/items(" + ID + ")/AttachmentFiles/add(FileName='" + file.name + "')",  
	            method: 'POST',  
	            async: false,  
	            data: buffer,  
	            processData: false,  
	            headers: {  
	                "Accept": "application/json; odata=verbose",  
	                "content-type": "application/json; odata=verbose", 	               
	                "X-RequestDigest": document.getElementById("__REQUESTDIGEST").value  
	
	            }, success: onAttachmentSucessDraft,error:OnAttachmentErrorDraft
	        }); 
	    });  
	    function onAttachmentSucessDraft(data) 
	    {  
	        fileUploadeCount++;	  
	        vStatusSuccFailed='Success';        
	        if (arraycount == fileUploadeCount) 
	        {  
	            console.log(data + ' uploaded successfully'); 
	            vStatusSuccFailed='Success';
				off();
				UpdateCounterDraft();				            
	        }  
	    } 
	    function OnAttachmentErrorDraft()
	    {
	      console.log('Error during file upload');
	      var vItemID=document.getElementById("hdnApprvlNotesItemID").value;
	      RemoveListItem(vItemID);	     	
	      off();
	      alert('Data could not be saved. Please refresh the browser and try again.');
	      window.location='/sites/teams/NFA/SitePages/NewNFA.aspx';
	    } 
	}  


function UpdateCounterDraft() 
 {
		debugger;
		var itemId = document.getElementById("hdnCounterItemID").value; 
		var hdnCounter= document.getElementById("hdnCounter").value;  
		hdnCounter=parseInt(hdnCounter)+1;
        $.ajax({
        
            url: _spPageContextInfo.webAbsoluteUrl +"/_api/web/lists/getByTitle('CounterMaster')/getItemByStringId('" +itemId + "')",
            type: "POST",
            contentType: "application/json;odata=verbose",
            async: false,
            data: JSON.stringify(
            {
                '__metadata': {
                    'type': 'SP.Data.CounterMasterListItem'
                },
                'Counter': hdnCounter.toString()              
            }),
            headers: {
                "accept": "application/json;odata=verbose",               
                "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                "IF-MATCH": "*",
                "X-Http-Method": "PATCH"
            },
            success: function (data) 
            {
                var vmsg="NFA "+document.getElementById("hdnDigitalRequestNo").value+" has been saved as draft.";     
                alert(vmsg);    
                window.location='https://brookfield.sharepoint.com/sites/teams/NFA';      
            },
            error: function (e) 
            {
            }
        });
}

function GetCounterDraft() 
{    
    debugger;
    var vRetVal='';
    var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('CounterMaster')//items?$top=500&$select=ID,Title,Counter,FunctionAbbr&$filter=(FunctionID eq '99')";
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) {
             debugger;

            var Logg = data.d.results;            
            try 
            {            
				var dtt = new Date();
				var mm = dtt.getMonth(); 
				var yyyy = dtt.getFullYear();
				var month = new Array();
				month[0] = "Jan";
				month[1] = "Feb";
				month[2] = "Mar";
				month[3] = "Apr";
				month[4] = "May";
				month[5] = "Jun";
				month[6] = "Jul";
				month[7] = "Aug";
				month[8] = "Sep";
				month[9] = "Oct";
				month[10] = "Nov";
				month[11] = "Dec";
				var mName = month[mm];
	            if(data.d.results.length > 0)
	            {	             
	             vRetVal="Draft"+'/'+yyyy+'/'+mName+'/'+data.d.results[0].Counter;
	             document.getElementById("hdnDigitalRequestNo").value=vRetVal;
	             document.getElementById("hdnCounter").value=data.d.results[0].Counter;
	             document.getElementById("hdnCounterItemID").value=data.d.results[0].ID;
	            }
	        }
            catch (e) {                               
            }
        },
        error: function (data) {
             debugger;
            console.log(data.responseText);
        }
    });  
    return vRetVal;
}


function ValidateDraftNFA()
{
    debugger;
    var msg = "";
    var focusId="";
    var rtvalue = true; 
    
    var ddlFunctionName  = $('#ddlFunctionName').val();
    var ddlProject       = $('#ddlProject').val();
    
  
    var vCount=0; var vFileSpclChar=0;
    if (ddlFunctionName== '' || ddlFunctionName=="0" || ddlFunctionName=="-Select Function-") 
    {  
        msg += "<font face='Calibri'><b>Please select Originator Department.</b></font>" + "<br/>";
        if (focusId== "") 
        {
         focusId=  $('#ddlFunctionName');
        }		
    }     

   
    if (ddlProject== '' || ddlProject == "0" || ddlProject == "-Select Project-") 
    {   
        msg += "<font face='Calibri'><b>Please select Project.</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#ddlProject');	
        } 	
    }
    
   
     var vSpecialChar=['~','!','#','$','%','^','&','*','(',')','<','>','{','}','|'];
     var vSizeOk=0;
     $("#attachFilesHolder input:file").each(function () 
     {   
        if ($(this)[0].files[0]) 
        {      
           var vfname= $(this)[0].files[0].name;
           var vFileSize=0;
           vFileSize=$(this)[0].files[0].size;
           if(parseFloat(vFileSize) > 5242880)
           {
            vSizeOk= parseInt(vSizeOk)+1;
           }
           if(includesMethod(vfname,vSpecialChar)==true)
           {
             vFileSpclChar=parseInt(vFileSpclChar)+1;
           }
           vCount=parseInt(vCount)+1;          
        }   
     });
     
     if(parseInt(vCount) > 0 && parseInt(vFileSpclChar) > 0)
     {
        msg += "<font face='Calibri'><b>One of supporting document name is having special characters .</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#file_input');	
        } 
     }
     else if(parseInt(vCount) > 0 && parseInt(vSizeOk) > 0)
     {
        msg += "<font face='Calibri'><b>A single file should not more than 5 MB .</b></font>" + "<br/>";        
        if (focusId== "") 
        {
          focusId=  $('#file_input');	
        } 

     }
    
    if (msg != "")
    {
      warningalert1(msg ,focusId);
      rtvalue = false;
	  return false;
    }   
    return rtvalue;
}




function UploadSoftCopy()
{    
    debugger;      
    CreateDocumentAttachment();
    if(parseInt(vAttachID) > 0)
    {    
    	openModelDialogPopup(vAttachID);
    }    
}

function openModelDialogPopup(vItemID) 
{      
    var vSoftCopyURL= "/sites/teams/NFA/Lists/NFADocs/UploadSCPTT.aspx?ID="+vItemID;
    var dialogOptions = 
    {  
        url: vSoftCopyURL,  
        title: 'Upload NFA Supporting Documents',  
        allowMaximize: false,  
        showClose: true,  
        width: 1000,  
        height: 450,
	    dialogReturnValueCallback: onCloseCallback
    };  
    SP.SOD.execute('sp.ui.dialog.js', 'SP.UI.ModalDialog.showModalDialog', dialogOptions);    
    //return false;  
}   

function onCloseCallback(dialogResult, returnValue) 
{ 
    getPreviousDocuments(vUniqNo);
}

function CreateDocumentAttachment()
{
        debugger;
        var vretval=false;
        if(vretval == false)
        {
	        var url = _spPageContextInfo.webAbsoluteUrl+ "/_api/web/lists/getbytitle('NFADocs')/items";
		    var data = JSON.stringify(  
		    {  
		        "__metadata":   
		        {  
		            "type": "SP.Data.NFADocsListItem"  
		        }, 
		        "Title":vUniqNo,
		        "UniqNo": vUniqNo,
			});  
		    $.ajax  
		    ({  
			        url: url,  
			        type: "POST",  
			        data: data,
			        async: false,  
			        headers: {  
			            "Accept": "application/json;odata=verbose",  
			            "Content-Type": "application/json;odata=verbose",  
			            "X-RequestDigest": $("#__REQUESTDIGEST").val()	    
			     },  
			     success: function(data)   
			  	 {  
			  	  vAttachID=data.d.Id;	  		  	 			    
			     },  
			     error: function(data)  
			     { 
			       console.log(data.responseText);
			     }  
	    	 });
          }
}


function getPreviousDocuments(vItemID) 
{ 
    var requestUri = _spPageContextInfo.webAbsoluteUrl+ "/_api/web/lists/getbytitle('NFADocs')//items?$select=Title,AttachmentFiles&$expand=AttachmentFiles&$filter=(UniqNo eq '"+vUniqNo+"')&$orderby=Created desc";
    debugger;
    var requestHeaders = { "accept": "application/json;odata=verbose" };
    $.ajax({
        url: requestUri,
        contentType: "application/json;odata=verbose",
        headers: requestHeaders,
        async: false,
        cache: false,
        success: function (data) 
        {
           debugger;
           var Logg = data.d.results; 
           if(Logg.length > 0)
           { 
                var vHtml="";
                for(var i=0; i<Logg.length;i++)
                {
	                var vAttch=data.d.results[i].AttachmentFiles.results;
	                if(vAttch.length>0)
	                {
		                 
		                 for(var j=0;j<vAttch.length; j++)
		                 {
		                   vHtml += "<a href='"+vAttch[j].ServerRelativeUrl+"' target='_blank'><font color='#000080' size='2'><b>"+vAttch[j].FileName+"</b></font></a><br/>";		                  
		                 }
		             }
	             } 
	             if(vHtml != "")
	             {
	             	$('#dvAttachment').html(vHtml);  
	             }            
		   }
        },
        error: function (data) {
            console.log(data.responseText);
        }
    });        
}


    