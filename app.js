// เมื่อ refresh browser ให้เคลียร์ file input ออกไปด้วย
window.addEventListener("load", () => {
  const fileInput   = document.querySelector("#fileInput");
  const fileNameDiv = document.getElementById("fileName");

  fileInput.value = "";          // ล้าง input[type="file"]
  fileNameDiv.textContent = "";  // ล้างชื่อไฟล์บนหน้า
});

const btnFetchReport = document.querySelector('#fetchReport')

btnFetchReport.addEventListener('click', async () => {
    // Get authentication
    const {domain, requestOptions} = getAuth()

    // Clear report 
    clearReport();

    // RUN FUNCTION: GET DATA FROM CSV FILE
    const fileInput   = document.querySelector("#fileInput");
    let dataFileCSV_arr = []
    if (fileInput.files[0] !== undefined) {
      // console.log(fileInput.files[0])
      dataFileCSV_arr = await getDataFromCSVInput(fileInput);
      // console.log(dataFileCSV)
    }
    

    // Input
    const ticketType = document.querySelector('#ticketType').value
    const ticketId   = document.querySelector('#ticketId').value

    // Validate ticket id input
    ticketId === '' | null ? alert('กรุณาใส่เลข ticket id') : null
    const hasTicketId = ticketId !== '' | null

    if (hasTicketId) {

    // ส่วนของ ticket object
    let   dataTicket = await getTicketById(ticketType, ticketId, domain, requestOptions);
    
    if (dataTicket !== null) {
      // Loading report 
      loadingReport();

      // Get ticket type in case of sr/inc ticket
      let ticketSubType = ''
     
      if(ticketType === 'tickets') {
        const ticket_subtype  = await dataTicket.type; 
        ticketSubType = ticket_subtype;
      }  

      const dataFormField = await getFormFields(ticketType, domain, requestOptions);

      // ส่วนของ Approval array
      let dataApprovals_arr = await getApprovalByTicketId(ticketType, ticketId, domain, requestOptions);

      // ส่วนของ Conversation array
      let dataConversation_arr = await getConversionByTicketId(ticketId, domain, requestOptions) !== null ? await getConversionByTicketId(ticketId, domain, requestOptions): [];
      

      // ส่วนของ Requested Item array (เฉพาะ service request)
      let dataRequestedItems_arr    = ticketSubType === "Service Request" ? await getRequestedItemsByTicketId(ticketId, domain, requestOptions) : []
      let serviceItemId_arr         = []
      let dataServiceCatalog_arr    = []
      if (await dataRequestedItems_arr.length > 0) {
        dataRequestedItems_arr.forEach((i, idx) => {
          const {
            service_item_id,
            stage
          } = i

          // เก็บ service item id ไว้ใน array
          serviceItemId_arr.push(service_item_id)

          // แปลงค่า state id เป็น value
          const stage_id = stage
          const stageIdMap = {
            1: "Requested",
            2: "Delivered",
            3: "Cancelled",
            4: "Fulfilled",
            5: "Partially Fulfilled"
          }
          let stage_value = stage_id !== null ? stageIdMap[stage_id] : null
          dataRequestedItems_arr[idx]['stage_value'] = stage_value
        })
        
        for (const [idx, i] of serviceItemId_arr.entries()) {
          const itemId = i
          let resultObj = await getServiceCatalogByItemId(itemId, domain, requestOptions)
          dataServiceCatalog_arr.push(resultObj)
        }
        console.log('requested items', dataRequestedItems_arr)
        console.log('service catalog', dataServiceCatalog_arr)
      }
      
      

      // ตัวแปรต่างๆ ที่ต้องการแปลงจาก id เป็นชื่อ
      const requester_id    = await dataTicket.requester_id 
      const agent_id        = await dataTicket.agent_id       !== undefined ? await dataTicket.agent_id : null;
      const department_id   = await dataTicket.department_id
      const group_id        = await dataTicket.group_id
      const status_id       = await dataTicket.status
      const source_id       = await dataTicket.source
      const urgency_id      = await dataTicket.urgency
      const impact_id       = await dataTicket.impact
      const risk_id         = await dataTicket.risk
      const priority_id     = await dataTicket.priority
      const change_type_id  = await dataTicket.change_type
      const workspace_id    = await dataTicket.workspace_id



      const dataRequester         = requester_id                       !== null ? await getRequesterById(requester_id, domain, requestOptions)                       : null;
      const dataAgent             = agent_id                           !== null ? await getAgentById(agent_id, domain, requestOptions)                               : null;
      const dataReportingManager  = dataRequester.reporting_manager_id !== null ? await getRequesterById(dataRequester.reporting_manager_id, domain, requestOptions) : null;
      const dataDepartment        = department_id                      !== null ? await getDepartmentById(department_id, domain, requestOptions)                     : null;
      const dataGroup             = group_id                           !== null ? await getGroupById(group_id, domain, requestOptions)                               : null;


      // จัดการ custom fields
      const custom_fields              = await dataTicket.custom_fields
      const customfields_keys_arr      = Object.keys(custom_fields);
      const customfields_values_arr    = Object.values(custom_fields);
      let customfields_name_arr        = [];
      let customfields_type_arr        = [];
      let customfields_value_trans_arr = [];
      let custom_fields_value          = [];

      // console.log(dataFormField)

      for (const [idx, i] of customfields_keys_arr.entries()) {

          let found = dataFormField.find(item => item.name === i);

          if (found) {

            let { label, field_type } = found

            customfields_name_arr .push(label)
            customfields_type_arr .push(field_type)

            if (field_type === "custom_lookup") {
              // customfields_value_trans_arr.push('ต้องเอา user id ไปหา =="')
              const customfield_requester_id    = customfields_values_arr[idx]
              const customfield_requester       = customfield_requester_id !== null ? await getRequesterById(customfield_requester_id, domain, requestOptions)        : null
              const customfield_requester_value = customfield_requester_id !== null ? await customfield_requester.first_name + " " + customfield_requester.last_name  : null
              customfields_value_trans_arr.push(customfield_requester_value)
            } else if (field_type === "custom_date") {
                const rawDate = customfields_values_arr[idx]
                const localDate = new Date(rawDate).toLocaleString("en-US", {
                  timeZone: "Asia/Bangkok", // GMT+7
                  weekday:  "short",
                  day:      "2-digit",
                  month:    "short",
                  hour:     "numeric",
                  minute:   "2-digit",
                  hour12:   true,
                })
                customfields_value_trans_arr.push(localDate)
            } else {
                customfields_value_trans_arr.push(customfields_values_arr[idx])
            }

          }  else {
            console.warn(`No match found for name = ${i}`)
          }

      }

      customfields_name_arr.forEach((i, idx) => {
          const object_name = customfields_name_arr[idx]
          const object_value = customfields_value_trans_arr[idx]

          if (customfields_type_arr[idx] === "custom_date") {
            
            custom_fields_value.push({[object_name]: object_value})
            // custom_fields_value[i] = customfields_value_trans_arr[idx]
            // สำหรับปรับชื่อ key ด้วย underscore และปรับเป็นตัวเล็ก และนำหน้าด้วย local_
            // const add_local_name = "local_" + i.toLowerCase().replace(/\s+/g, "_") 
            // custom_fields_value[add_local_name] = customfields_value_trans_arr[idx]
          } else {
            custom_fields_value.push({[object_name]: object_value})
            // custom_fields_value[i] = customfields_value_trans_arr[idx]
            // สำหรับปรับชื่อ key ด้วย underscore และปรับเป็นตัวเล็ก
            // custom_fields_value[i.toLowerCase().replace(/\s+/g, "_")] = customfields_value_trans_arr[idx] 
          }
      })


      // console.log(customfields_keys_arr)
      // console.log(customfields_values_arr)
      // console.log(customfields_name_arr)
      // console.log(customfields_type_arr)
      // console.log(customfields_value_trans_arr)
      // console.log(custom_fields_value)
      
      // หาค่า value จาก id
      const agent_value              = agent_id     !== null ? await dataAgent.first_name     + " " + dataAgent.last_name     : null
      const requester_value          = requester_id !== null ? await dataRequester.first_name + " " + dataRequester.last_name : null
      const requester_mail           = await dataRequester.primary_email
      const reportingmanager_value   = dataReportingManager !== null ? await dataReportingManager.first_name + " " + dataReportingManager.last_name : null
      const status_value             = await dataFormField.find(item => item.name === "status")?.choices.find(item => item.id === status_id)           ?.value
      const source_value             = await dataFormField.find(item => item.name === "source")?.choices.find(item => item.id === source_id)           ?.value
      const urgency_value            = await dataFormField.find(item => item.name === "urgency")?.choices.find(item => item.id === urgency_id)         ?.value
      const department_value         = department_id !== null ? await dataDepartment.name : null
      const group_value              = dataGroup     !== null ? await dataGroup.name      : null
      const impact_value             = await dataFormField.find(item => item.name === "impact")      ?.choices.find(item => item.id === impact_id)     ?.value
      const risk_value               = await dataFormField.find(item => item.name === "risk")        ?.choices.find(item => item.id === risk_id)       ?.value
      const priority_value           = await dataFormField.find(item => item.name === "priority")    ?.choices.find(item => item.id === priority_id)   ?.value
      const change_type_value        = await dataFormField.find(item => item.name === "change_type") ?.choices.find(item => item.id === change_type_id)?.value
      const workspace_value          = await dataFormField.find(item => item.name === "workspace_id")?.choices.find(item => item.id === workspace_id)  ?.value


      // convert utc to local time and format to 'Thu, 10 Apr 1:51 PM'
      // created_at, updated_at, planned_start_date, planned_end_date, due_by
      const local_created_at          = await dataTicket.created_at         !== null ? convertLocalFormat(dataTicket.created_at)         : null
      const local_updated_at          = await dataTicket.updated_at         !== null ? convertLocalFormat(dataTicket.updated_at)         : null
      const local_planned_start_date  = await dataTicket.planned_start_date !== null ? convertLocalFormat(dataTicket.planned_start_date) : null
      const local_planned_end_date    = await dataTicket.planned_end_date   !== null ? convertLocalFormat(dataTicket.planned_end_date)   : null
      const local_due_by              = await dataTicket.due_by             !== null ? convertLocalFormat(dataTicket.due_by)             : null

      // เพิ่ม value ที่หาค่ามาจาก id เข้า ticket object
      dataTicket['agent_value']              = agent_value
      dataTicket['requester_value']          = requester_value
      dataTicket['requester_mail']           = requester_mail
      dataTicket['reportingmanager_value']   = reportingmanager_value
      dataTicket['status_value']             = status_value
      dataTicket['source_value']             = source_value
      dataTicket['urgency_value']            = urgency_value
      dataTicket['department_value']         = department_value
      dataTicket['impact_value']             = impact_value
      dataTicket['risk_value']               = risk_value
      dataTicket['priority_value']           = priority_value
      dataTicket['change_type_value']        = change_type_value
      dataTicket['workspace_value']          = workspace_value
      dataTicket['group_value']              = group_value
      dataTicket['custom_fields_value']      = custom_fields_value
      dataTicket['local_created_at']         = local_created_at
      dataTicket['local_updated_at']         = local_updated_at
      dataTicket['local_planned_start_date'] = local_planned_start_date
      dataTicket['local_planned_end_date']   = local_planned_end_date
      dataTicket['local_due_by']             = local_due_by
      
      console.log('ข้อมูล ticket', dataTicket)


      // loop หา value จาก id ของ approver array
      if (dataApprovals_arr !== null) {
        for (const [idx, i] of dataApprovals_arr.entries()) {
          const {
            approval_type,
            updated_at
          } = i
      
          const approvalTypeMap = {
            1: "Following has to approve", //"To be approved by Everyone",
            2: "To be approved by Anyone",
            3: "To be approved by Majority",
            4: "To be approved by First Responder"
          };
          const approval_type_value       = approval_type !== null ? approvalTypeMap[approval_type] || "Unknown Approval Type" : null;
          const local_approval_updated_at = updated_at    !== null ? convertLocalFormat(updated_at)                            : null

          // console.log(approval_type_value)
          // console.log(local_approval_updated_at)

          // เพิ่ม value เข้าแต่ละ approver object ใน array
          dataApprovals_arr[idx]['approval_type_value'] = approval_type_value
          dataApprovals_arr[idx]['local_approval_updated_at'] = local_approval_updated_at
        }
      }
      
      console.log('ข้อมูล approvals', dataApprovals_arr)




      // loop หา value จาก id ของ conversation array
      if (dataConversation_arr.length > 0) {
        for (const [idx, i] of dataConversation_arr.entries()) {
          const {
            source,
            created_at,
            user_id
          } = i
      
          // หา ชื่อนามสกุล และ email จาก user_id
        const {first_name, last_name, primary_email} = await getRequesterById(user_id, domain, requestOptions)
        const from_name = primary_email !== null ? first_name + ' ' + last_name : null

        const sourceTypeMap = {
            0: "email",
            1: "form",
            2: "note",
            3: "status",
            4: "meta",
            5: "feedback",
            6: "forward_email"
          };

          const source_value = source !== null ? sourceTypeMap[source] : null
          const local_conversation_created_at = created_at !== null ? convertLocalFormat(created_at) : null

          // เพิ่ม value เข้าแต่ละ conversation object ใน array
          dataConversation_arr[idx]['from_name'] = from_name;
          dataConversation_arr[idx]['primary_email'] = primary_email;
          dataConversation_arr[idx]['source_value'] = source_value;
          dataConversation_arr[idx]['local_conversation_created_at'] = local_conversation_created_at;
        }
      }
      
      console.log('ข้อมูล conversations', dataConversation_arr)

      // สร้างรายงาน
      generateReport(
        ticketType, 
        ticketSubType, 
        dataTicket, 
        dataApprovals_arr, 
        dataConversation_arr, 
        dataRequestedItems_arr, 
        dataServiceCatalog_arr, 
        dataFileCSV_arr
      )
      }
    }
})

// GET AUTHENCATION
function getAuth() {
  
  const domain = document.querySelector('#domain').value;

  const username = document.querySelector('#authToken').value;
  const password = "X";

  // สร้าง string "user:password"
  const credentials = `${username}:${password}`;

  // แปลงเป็น Base64
  const encodedCredentials = btoa(credentials);
  
  // สร้าง Headers object
  const myHeaders = new Headers();
  myHeaders.append("Authorization", `Basic ${encodedCredentials}`);

  const requestOptions = {
    method: "GET",
    headers: myHeaders,
    redirect: "follow",
  };

  return {
    domain: domain,
    requestOptions: requestOptions
  };
}

// API GET TICKET DETAILS FOLLOWS BY TICKET TYPE AND TICKET ID
async function getTicketById(ticket_type, ticket_id, domain, requestOptions) {

  const URL_GETTICKET = `https://${domain}.freshservice.com/api/v2/${ticket_type}/${ticket_id}`;

  try {
    const response = await fetch(URL_GETTICKET, requestOptions);
    const result = await response.json();

    if (response.ok) {
      // ใช้ mapping แทน switch
      const typeMap = {
        tickets: result.ticket,
        changes: result.change
      };
      return typeMap[ticket_type] ?? null;
    } else {
      alert(result.message || `Error ${response.status}`);
      return null;
    }

  } catch (e) {
    console.error(e);
    alert(`ไม่พบข้อมูล ticket id ${ticket_id}`);
    return null;
  }
}

// API GET APPROVAL DETAIL
async function getApprovalByTicketId(ticket_type, ticket_id, domain, requestOptions) {
  const URL_GETAPPROVALS = `https://${domain}.freshservice.com/api/v2/${ticket_type}/${ticket_id}/approvals`

  try {
    const response = await fetch(URL_GETAPPROVALS, requestOptions);
    const result   = await response.json();
    const dataObj  = await result.approvals;
    return dataObj;
  } catch (e) {
    //alert(`ไม่พบข้อมูล approval`)
    return null
  }
}

// API GET CONVERSATION (COMMENT) DETAIL
async function getConversionByTicketId(ticket_id, domain, requestOptions) {
  const URL_GETCONVERSACTION = `https://${domain}.freshservice.com/api/v2/tickets/${ticket_id}/conversations`

  try {
    const response = await fetch(URL_GETCONVERSACTION, requestOptions);
    const result   = await response.json();
    const dataObj  = await result.conversations;
    return dataObj;
  } catch (e) {
    // alert(`ไม่พบข้อมูล conversion`)
    return null
  }
}

// API GET REQUESTER DETAIL
async function getRequesterById(requester_id, domain, requestOptions) {
  const URL_GETREQUESTER = `https://${domain}.freshservice.com/api/v2/requesters/${requester_id}`

  try {
    const response = await fetch(URL_GETREQUESTER, requestOptions);
    const result   = await response.json();
    const dataObj  = await result.requester;
    return dataObj;
  } catch (e) {
    alert(`ไม่พบข้อมูล requester id ${requester_id}`)
  }
}

// API GET AGENT DETAIL
async function getAgentById(agent_id, domain, requestOptions) {
  const URL_GETAGENT = `https://${domain}.freshservice.com/api/v2/agents/${agent_id}`

  try {
    const response = await fetch(URL_GETAGENT, requestOptions);
    const result   = await response.json();
    const dataObj  = await result.agent;
    return dataObj;
  } catch (e) {
    alert(`ไม่พบข้อมูล agent id ${agent_id}`)
  }
}

// API GET DEPARTMENT DETAIL
async function getDepartmentById(department_id, domain, requestOptions) {
  const URL_GETDEPARTMENT = `https://${domain}.freshservice.com/api/v2/departments/${department_id}`

  try {
    const response = await fetch(URL_GETDEPARTMENT, requestOptions);
    const result   = await response.json();
    const dataObj  = await result.department;
    return dataObj;
  } catch (e) {
    alert(`ไม่พบข้อมูล department id ${department_id}`)
  }
}

// API GET GROUP DETAIL
async function getGroupById(group_id, domain, requestOptions) {
  const URL_GETGROUP = `https://${domain}.freshservice.com/api/v2/groups/${group_id}`

  try {
    const response = await fetch(URL_GETGROUP, requestOptions);
    const result   = await response.json();
    const dataObj  = await result.group;
    return dataObj;
  } catch (e) {
    alert(`ไม่พบข้อมูล group id ${group_id}`)
  }
}


// API GET REQUESTED ITEMS DETAIL
async function getRequestedItemsByTicketId(ticket_id, domain, requestOptions) {
  const URL_GETREQUESTEDITEMS = `https://${domain}.freshservice.com/api/v2/tickets/${ticket_id}/requested_items`

  try {
    const response = await fetch(URL_GETREQUESTEDITEMS, requestOptions);
    const result   = await response.json();
    const dataObj  = await result.requested_items;
    return dataObj;
  } catch (e) {
    alert(`ไม่พบข้อมูล requester items ของ ticket id ${ticket_id}`)
  }
}

// API GET SERVICE CATALOG DETAIL
async function getServiceCatalogByItemId(item_id, domain, requestOptions) {
  const URL_GETSERVICECATALOG = `https://${domain}.freshservice.com/api/v2/service_catalog/items/${item_id}`

  try {
    const response = await fetch(URL_GETSERVICECATALOG, requestOptions);
    const result   = await response.json();
    const dataObj  = await result.service_item;
    return dataObj;
  } catch (e) {
    alert(`ไม่พบข้อมูล service catalog ของ item id ${item_id}`)
  }
}


// API GET FORM FIELD DETAIL
async function getFormFields(ticket_type, domain, requestOptions) {
  let formFieldName = '';

  switch (ticket_type) {
    case "tickets": 
      formFieldName = "ticket_form_fields";
      break;
    case "changes": 
      formFieldName = "change_form_fields";
      break;
    default:
      throw new Error(`Unknown ticket_type: ${ticket_type}`);
  }

  const URL_GETFORMFIELD = `https://${domain}.freshservice.com/api/v2/${formFieldName}?per_page=100`

  try {
    const response = await fetch(URL_GETFORMFIELD, requestOptions);
    const result   = await response.json();
    const dataObj  = await result;

    switch(ticket_type) {
        case "tickets": return dataObj.ticket_fields
        case "changes": return dataObj.change_fields
    }
  } catch (e) {
    alert(`ไม่พบข้อมูล form field`)
  }
}

// CONVERT TIME TO LOCAL AND FORMAT DATETIME
function convertLocalFormat(utcDate) {
  // Convert to Date object
  const date = new Date(utcDate);

  // Format with Intl.DateTimeFormat in GMT+7
  const formatter = new Intl.DateTimeFormat("en-US", {
    timeZone: "Asia/Bangkok", // GMT+7
    weekday:  "short",
    day:      "2-digit",
    month:    "short",
    hour:     "numeric",
    minute:   "2-digit",
    hour12:   true,
  });

  const formatted = formatter.format(date);
  return formatted
}

// CONVERT STRING TO PROPER CASE
function toProperCase(str) {
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}

// CLEAR CUSTOM REPORT
function clearReport() {
  document.querySelector('.ctn-report').innerHTML = '<div></div>'
}

// LOADING CUSTOM REPORT
function loadingReport() {
  document.querySelector('.ctn-report').innerHTML = '<div>Report is loading ...</div>'
}

// GENERATE CUSTOM REPORT
async function generateReport(
  ticket_type, 
  ticket_subtype, 
  data_ticket, 
  data_approval_arr,
  data_conversation_arr,
  data_requesteditems_arr, 
  data_servicecatalog_arr, 
  data_activitylog_arr
) {
    let report = document.querySelector('.ctn-report')

    // activity log จากไฟล์แนบ
    let activitylog_html = ''
    if (data_activitylog_arr.length > 0 ){
      activitylog_html += '<div class="saparate-line"></div>'
      activitylog_html += '<h3>ACTIVITIES LOG</h3>'
      activitylog_html += csvJsonToTable(data_activitylog_arr)
    }

    // ส่วน custom fields (array) ซึ่งเป็นส่วนย่อยข้างในของ ticket properties
      const custom_fields_value  = data_ticket.custom_fields_value
      let customfields_html = ''
      if (custom_fields_value.length > 0) {
        custom_fields_value.forEach((obj) => {
        const key = Object.keys(obj)[0]     // ดึง key
        const value = obj[key]              // ดึง value ตาม key นั้น

        customfields_html += `
          <div class='tp-grid-item'>
            <div class='item-topic'>${key}</div>
            <div class='item-value'>${value || '--'}</div>
          </div>
        `
        })
      }

    // ส่วนของ approval array
    let approvallog_html = ''

    if (data_approval_arr !== null) {
    
      approvallog_html += `
        <div class='saparate-line'></div>
        <h3>APPROVAL LOGS</h3>
      `
      data_approval_arr.forEach((i, idx) => {
        const {
          approval_type_value,
          approver_name,
          approval_status,
          local_approval_updated_at,
          latest_remark
        } = i

        const approval_status_name = approval_status.name === 'peer_responded' ? `Already approved by ${data_approval_arr[idx - 1].approver_name}` : approval_status.name
        const approver_status_propercase = toProperCase(approval_status_name)

        
        approvallog_html += `<div class='item-topic'>${approval_type_value}</div>`
        approvallog_html += `
          <div class='ap-grid'>
            <div>${approver_name}</div>
              <div>
                <div>${approver_status_propercase} on ${local_approval_updated_at}</div>
                <div>${latest_remark}</div>
              </div>
            </div>
          </div>
        `
      })
    }


    // ส่วนของ conversaton array
    let conversation_html = ''

    if (data_conversation_arr.length > 0) {
    
      conversation_html += `
        <div class='saparate-line'></div>
        <h3>COMMENTS</h3>
      `
      for (let idx = data_conversation_arr.length - 1; idx >= 0; idx--) {
        const {
          from_email,
          from_name,
          primary_email,
          incoming,
          source_value,
          local_conversation_created_at,
          body
        } = data_conversation_arr[idx]        

        conversation_html += `
          <div>From <b>${from_name}</b> (${primary_email}) on <b>${local_conversation_created_at} as ${incoming ? 'Inbound': 'Outbound'} ${source_value}</b></div>
          ${body}
        `
      }

    }
    
    // report สำหรับ tickets
    if (ticket_type === 'tickets') {
      
      // ส่วนของ ticket
      let {
          subject,
          id, // ticket id
          requester_value,
          requester_mail,
          local_created_at,
          workspace_value,
          status_value,
          priority_value,
          source_value,
          type,
          impact_value,
          urgency_value,
          group_value,
          agent_value,
          category,
          local_planned_start_date,
          local_planned_end_date,
          department_value,
          planned_effort,
          custom_fields_value, // เป็น array
          description,
          local_due_by,
          //tags ไม่รู้หาจากไหน
          resolution_notes_html
      } = data_ticket;

      // ส่วนของ service items (กรณีที่เป็น service request)
      const number_items = data_requesteditems_arr.length
      let serviceitems_customfields_html = ''
      let additional_notes_html = ''
      let serviceitems_html     = ''

      if (number_items > 0) {
        // เขียนวนลูป ...
        //custom_paragraphs จาก requested item
        data_requesteditems_arr.forEach((i, idx) => {
          const {
            service_item_name,
            stage_value,
            custom_fields: requesteditem_custom_fields,
            custom_paragraphs,
            ff_single_line_tf
          } = i

          const {
            description,
            custom_fields: catalog_custom_fields
          } = data_servicecatalog_arr[idx]



          ff_single_line_tf.forEach((i, idx) => {
            const requesteditem_custom_fields_topic = Object.keys(requesteditem_custom_fields).filter(k => requesteditem_custom_fields[k] === i)[0]
            // ค้นหา label จาก name ที่ได้จาก custom field ของ requested item
            let catalog_custom_fields_name_arr = []
            let catalog_custom_fields_label_arr = []
            catalog_custom_fields.map((c) => {
              catalog_custom_fields_name_arr.push(c.name)
              catalog_custom_fields_label_arr.push(c.label)
            })
            
            // console.log(requesteditem_custom_fields_topic)
            // console.log(catalog_custom_fields_name_arr)
            // console.log(catalog_custom_fields_label_arr)

            const name_index = catalog_custom_fields_name_arr.indexOf(requesteditem_custom_fields_topic)
            const custom_fields_label = catalog_custom_fields_label_arr[name_index]

            // console.log(name_index)

            serviceitems_customfields_html += `
              <div class='si-grid-item'>
                <div class='item-topic'>${custom_fields_label}</div>
                <div class='item-value'>${i}</div>
              </div>
            `
          })

          
          custom_paragraphs.forEach((i, idx) => {
            additional_notes_html += `
              <div class='si-grid-item'>
                <div class='item-topic'>Additional Note (${idx + 1})</div>
                <div class='item-value'>${i}</div>
              </div>
            `
          })
          
          


          // ส่วนย่อย service item ของ report
          serviceitems_html += `
            <h5>${service_item_name}</h5>
            <h5>DESCRIPTION</h5>
            ${description}

            <div class='si-grid'>
            
              <div class='si-grid-item'>
                <div class='item-topic'>Stage</div>
                <div class='item-value'>${stage_value}</div>
              </div>

              ${serviceitems_customfields_html}

              ${additional_notes_html}

            </div>
          `
        })
      }


      // ออก report
      report.innerHTML = `
          <h3>${subject}     <span>#${ticket_subtype === "Incident" ? "INC": "SR"}-${id}</span></h3>
          <div>by  <b>${requester_value}</b> (${requester_mail}) on <b>${local_created_at}</b> via <b>${source_value}</b></div>

          <div class='saparate-line'></div>

          <h3>TICKET PROPERTIES</h3>
          <div class='tp-grid'>
            
            <div class='tp-grid-item'>
              <div class='item-topic'>Workspace</div>
              <div class='item-value'>${workspace_value}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Priority</div>
              <div class='item-value'>${priority_value}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Status</div>
              <div class='item-value'>${status_value}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Source</div>
              <div class='item-value'>${source_value}</div>
            </div>

            
            <div class='tp-grid-item'>
              <div class='item-topic'>Type</div>
              <div class='item-value'>${type}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Urgency</div>
              <div class='item-value'>${urgency_value}</div>
            </div>
             <div class='tp-grid-item'>
              <div class='item-topic'>Impact</div>
              <div class='item-value'>${impact_value}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Group</div>
              <div class='item-value'>${group_value || '--'}</div>
            </div>

            <div class='tp-grid-item'>
              <div class='item-topic'>Agent</div>
              <div class='item-value'>${agent_value || '--'}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Department</div>
              <div class='item-value'>${department_value}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Category</div>
              <div class='item-value'>${category || '--'}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Planned Start Date</div>
              <div class='item-value'>${local_planned_start_date || '--'}</div>
            </div>

            <div class='tp-grid-item'>
              <div class='item-topic'>Planned End Date</div>
              <div class='item-value'>${local_planned_end_date || '--'}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Planned Effort</div>
              <div class='item-value'>${planned_effort || '--'}</div>
            </div>

            ${customfields_html}
            
            <div class='tp-grid-item'>
              <div class='item-topic'>Due by</div>
              <div class='item-value'>${local_due_by || '--'}</div>
            </div>

          </div>

          <div class='saparate-line'></div>
          
          <h3>DESCRIPTION</h3>
          <div>${description}</div>

          

          ${
            ticket_subtype === "Service Request"
            ? `
              <div class='saparate-line'></div>
              <h3>REQUESTED ITEMS (${number_items})</h3>
              ${serviceitems_html}              
              `
            : ''
          }

          
          ${approvallog_html || ''}
          
         
          ${`
            <div class='saparate-line'></div>
            <h3>RESOLUTION NOTE</h3> 
            ${resolution_notes_html || ''}
          `
          }
          
          ${conversation_html}

          ${activitylog_html}
      `
    }



    // report สำหรับ changes
    if (ticket_type === 'changes') {
      // ส่วนของ ticket
      let {
          subject,
          id, // ticket id
          requester_value,
          requester_mail,
          local_created_at,
          workspace_value,
          change_type_value,
          status_value,
          priority_value,
          impact_value,
          risk_value,
          group_value,
          agent_value,
          category,
          sub_category,
          item_category, 
          local_planned_start_date,
          local_planned_end_date,
          department_value,
          planned_effort,
          custom_fields_value,
          description,
          planning_fields
      } = data_ticket;

      // planning fields
      let {
        backout_plan,
        change_impact,
        reason_for_change,
        rollout_plan
      } = planning_fields;

      report.innerHTML = `
          <h3>${subject}     <span>#CHN-${id}</span></h3>
          <div>by  <b>${requester_value}</b> (${requester_mail}) on <b>${local_created_at}</b></div>

          <div class='saparate-line'></div>

          <h3>TICKET PROPERTIES</h3>
          <div class='tp-grid'>
            
            <div class='tp-grid-item'>
              <div class='item-topic'>Workspace</div>
              <div class='item-value'>${workspace_value}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Change Type</div>
              <div class='item-value'>${change_type_value || '--'}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Status</div>
              <div class='item-value'>${status_value}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Priority</div>
              <div class='item-value'>${priority_value}</div>
            </div>

            
            <div class='tp-grid-item'>
              <div class='item-topic'>Impact</div>
              <div class='item-value'>${impact_value}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Risk</div>
              <div class='item-value'>${risk_value || '--'}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Group</div>
              <div class='item-value'>${group_value || '--'}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Agent</div>
              <div class='item-value'>${agent_value || '--'} </div>
            </div>

            
            <div class='tp-grid-item'>
              <div class='item-topic'>Category</div>
              <div class='item-value'>${category || '--'}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Sub-Category</div>
              <div class='item-value'>${sub_category || '--'}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Item</div>
              <div class='item-value'>${item_category || '--'}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Planned Start Date</div>
              <div class='item-value'>${local_planned_start_date || '--'}</div>
            </div>

            
            <div class='tp-grid-item'>
              <div class='item-topic'>Planned End Date</div>
              <div class='item-value'>${local_planned_end_date || '--'}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Department</div>
              <div class='item-value'>${department_value}</div>
            </div>
            <div class='tp-grid-item'>
              <div class='item-topic'>Planned Effort</div>
              <div class='item-value'>${planned_effort || '--'}</div>
            </div>
            
            ${customfields_html}

            </div>
          

          <div class='saparate-line'></div>
          
          <h3>DESCRIPTION</h3>
          <div>${description}</div>

          <div class='saparate-line'></div>
          
          <h3>PLANNING</h3>
          <div class='pn-grid'>

            <div class='pn-grid-item'>
              <div class='item-topic'>Reason for Change</div>
              <div class='item-value'>${reason_for_change?.description_html || '--'}</div>
            </div>

            <div class='pn-grid-item'>
              <div class='item-topic'>Impact</div>
              <div class='item-value'>${change_impact?.description_html || '--'}</div>
            </div>

            <div class='pn-grid-item'>
              <div class='item-topic'>Rollout Plan</div>
              <div class='item-value'>${rollout_plan?.description_html || '--'}</div>
            </div>

            <div class='pn-grid-item'>
              <div class='item-topic'>Backout Plan</div>
              <div class='item-value'>${backout_plan?.description_html || '--'}</div>
            </div>

          </div>          

          ${approvallog_html}

          ${activitylog_html}

      `
    }
}






// จัดการไฟล์แนบ
const dropZone     = document.getElementById("dropZone");
const fileInput    = document.getElementById("fileInput");
const fileNameDiv  = document.getElementById("fileName");
const btnClearFile = document.getElementById("clearFile");

// เคลียร์ไฟล์
// ล้างไฟล์
btnClearFile.addEventListener("click", () => {
  fileInput.value = "";           // ล้าง input[type="file"]
  fileNameDiv.textContent = "";   // ล้างชื่อไฟล์บนหน้า
});



// คลิกพื้นที่ = เปิด file dialog
dropZone.addEventListener("click", () => fileInput.click());

// เลือกไฟล์จาก dialog
fileInput.addEventListener("change", (e) => handleFile(e.target.files[0]));

// Drag over
dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("dragover");
});

// Drag leave
dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});

// Drop
dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");
  const file = e.dataTransfer.files[0];
  handleFile(file);
});

// ฟังก์ชันตรวจสอบ CSV และแสดงไฟล์
function handleFile(file) {
  if (!file) return;

  // ตรวจสอบนามสกุล CSV
  if (file.type !== "text/csv" && !file.name.endsWith(".csv")) {
    alert("กรุณาเลือกไฟล์ CSV เท่านั้น");
    return;
  }

  fileNameDiv.textContent = `ไฟล์ที่เลือก: ${file.name} (${Math.round(file.size / 1024)} KB)`;

}

// GET DATA FROM CSV FILE
function getDataFromCSVInput(inputElement) {
  return new Promise((resolve, reject) => {
    const file = inputElement.files[0];
    if (!file) {
      resolve([]);
      return;
    }
    const reader = new FileReader();
    reader.onload = function (e) {
      const csvText = e.target.result;
      const cleanedCSV = removeNewlinesInQuotes(csvText);
      resolve(parseCSV(cleanedCSV));
    };
    reader.onerror = function (e) {
      reject(e);
    };
    reader.readAsText(file);
  });
}

// REMOVE NEW LINE IN QUOTES
function removeNewlinesInQuotes(csvText) {
  let inQuotes = false;
  let result = '';
  for (let i = 0; i < csvText.length; i++) {
    const char = csvText[i];
    if (char === '"') {
      inQuotes = !inQuotes;
      result += char;
    } else if ((char === '\n' || char === '\r') && inQuotes) {
      // skip newlines inside quotes
      continue;
    } else {
      result += char;
    }
  }
  return result;
}

// PARSE CSV
function parseCSV(csvText) {
  // Each record is always 4 columns: Name, Time, Content, Subcontent
  const lines = csvText.trim().split('\n').filter(line => line.trim() !== '');
  const result = [];
  for (let i = 1; i < lines.length; i++) { // skip header
    let line = lines[i];
    let columns = [];
    let current = '';
    let inQuotes = false;
    for (let j = 0; j < line.length; j++) {
      const char = line[j];
      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === ',' && !inQuotes) {
        columns.push(current.trim().replace(/^"|"$/g, ''));
        current = '';
      } else {
        current += char;
      }
    }
    columns.push(current.trim().replace(/^"|"$/g, ''));
    // Always return 4 columns, fill missing with ''
    result.push({
      Name:       columns[0] || '',
      Time:       columns[1] || '',
      Content:    columns[2] || '',
      Subcontent: columns[3] || ''
    });
  }
  return result;
}

// GENERATE TABLE FROM CSV FILE
function csvJsonToTable(arr) {
  if (!arr || arr.length === 0) return "<div>No data</div>";

  // Get headers from the first object
  const headers = Object.keys(arr[0]);
  let html = `
    <table>
      <thead>
        <tr>
  `;

  // Header row
  headers.forEach(header => {
    html += `<th class="csv-header">${header}</th>`;
  });

  html += `
        </tr>
      </thead>
    <tbody>
  `;

  // Data rows (reverse order: last line to first line)
  for (let i = arr.length - 1; i >= 0; i--) {
    const row = arr[i];
    // Highlight row if Content includes 'set Status as'
    const highlight = row.Content && row.Content.includes('set Status as') ? ' style="background:#ffe066;font-weight:bold;"' : '';
    html += `<tr${highlight}>`;
    headers.forEach(header => {
      html += `<td>${row[header]}</td>`;
    });
    html += '</tr>';
  }

  html += '</tbody></table>';
  return html;
}





// ปุ่ม Print
const btnPrintReport = document.querySelector('#printReport')
btnPrintReport.addEventListener('click', () => {
   printReport()
})

// PRINT FUNCTION
function printReport() {
  const ticketId   = document.querySelector('#ticketId').value

  const fileName = `report_ticket_number_${ticketId}`
  const content = document.querySelector(".ctn-report").innerHTML;
  const iframe = document.createElement('iframe');
  iframe.style.position = 'fixed';
  iframe.style.right    = '0';
  iframe.style.bottom   = '0';
  iframe.style.width    = '0';
  iframe.style.height   = '0';
  iframe.style.border   = '0';
  document.body.appendChild(iframe);

  const doc = iframe.contentWindow.document;
  doc.open();
  doc.write(`
    <html>
      <head>
        <title>${fileName}</title> <!-- กำหนดชื่อไฟล์เวลา Save as PDF -->
        <!-- โหลด style.css -->
        <link rel="stylesheet" type="text/css" href="style.css">
        <style>
          @page { size: A4; margin: 10mm; }
        </style>
      </head>
      <body>${content}</body>
    </html>
  `);
  doc.close();

  iframe.onload = () => {
    iframe.contentWindow.focus();
    iframe.contentWindow.print();
    setTimeout(() => document.body.removeChild(iframe), 100);
  };
}