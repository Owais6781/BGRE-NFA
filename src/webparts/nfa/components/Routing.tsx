import * as React from "react";
import { Routes, Route, HashRouter, Navigate } from "react-router-dom";
import { INfaProps } from "./INfaProps";
import "../../../Share/styles/global.scss";
import "primereact/resources/themes/lara-light-blue/theme.css";
import "primereact/resources/primereact.min.css";

import Sidebar from "./SideBar/Sidebar";
import Header from "./Layouts/Header";
import Home from "./ApprovalNotes/Home/Home";
import NewNFA from "./ApprovalNotes/NewNFA/NewNFA";
import ViewNFA from "./ApprovalNotes/ViewNFA/ViewNFA";
import Inbox from "./ApprovalNotes/Inbox/Inbox";
import ApproverPage from "./ApprovalNotes/ApproverPage/ApproverPage";
import QueryRespond from "./ApprovalNotes/QueryRespond/QueryRespond";
import ApprovalDetailsView from "./ApprovalNotes/ApprovalDetailsView/ApprovalDetailsView";
import ApprovedRejectedByMe from "./ApprovalNotes/ApprovedRejectedByMe/ApprovedRejectedByMe";
import DetailView from "./AdminSection/DetailView/DetailView";
import Masters from "./AdminSection/Masters/Masters";
import FunctionWiseStatus from "./Reports/FunctionWiseStatus/FunctionWiseStatus";
import ApproverWisePendency from "./Reports/ApproverWisePendency/ApproverWisePendency";
import AgeingReport from "./Reports/AgeingReport/AgeingReport";
import DayWiseAgeing from "./Reports/DayWiseAgeing/DayWiseAgeing";
import GenericReport from "./Reports/GenericReport/GenericReport";

const Routing: React.FC<INfaProps> = ({ context, sp }) => {
  return (
    <HashRouter>
      <div className="w-full flex" style={{ background: '#f1f5f9', padding: '10px', minHeight: 'calc(100vh - 140px)' }}>
        <Sidebar context={context} sp={sp} />

        <div className="flex flex-col flex-1 min-w-0">
          <Header context={context} />

          <main className="flex-1 w-full min-w-0" style={{ padding: '10px 0 30px 20px' }}>
            <Routes>
                <Route path="/"                    element={<Home context={context} sp={sp}/>} />
                <Route path="/new-nfa"             element={<NewNFA context={context} sp={sp}/>} />
                <Route path="/new-nfa/:id"         element={<NewNFA context={context} sp={sp}/>} />
                <Route path="/view-nfa/:id"        element={<ViewNFA context={context} sp={sp}/>} />
                <Route path="/approver-page/:id"   element={<ApproverPage context={context} sp={sp}/>} />
                <Route path="/query-respond/:id"   element={<QueryRespond context={context} sp={sp}/>} />
                <Route path="/inbox"               element={<Inbox context={context} sp={sp}/>} />
                <Route path="/approver-detail"     element={<ApprovalDetailsView context={context} sp={sp}/>} />
                <Route path="/approved-rejected"   element={<ApprovedRejectedByMe context={context} sp={sp}/>} />
                <Route path="/detail-view"         element={<DetailView />} />
                <Route path="/masters"             element={<Masters />} />
                <Route path="/function-wise-status"    element={<FunctionWiseStatus context={context} sp={sp}/>} />
                <Route path="/approver-wise-pendency"  element={<ApproverWisePendency />} />
                <Route path="/ageing-report"       element={<AgeingReport context={context} sp={sp}/>} />
                <Route path="/day-wise-ageing"     element={<DayWiseAgeing />} />
                <Route path="/generic-report"      element={<GenericReport />} />
                {/* Fallback to Home */}
                <Route path="*" element={<Navigate to="/" replace />} />
              </Routes>
            </main>
          </div>

        </div>
      
    </HashRouter>
  );
};

export default Routing;
