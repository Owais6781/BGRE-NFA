import * as React from "react";
import { Routes, Route, HashRouter } from "react-router-dom";
import { INfaProps } from "./INfaProps";
import "../../../Share/styles/global.scss";
import "primereact/resources/themes/lara-light-blue/theme.css";
import "primereact/resources/primereact.min.css";
import Sidebar from "./SideBar/Sidebar";
import Header from "./Layouts/Header";
import Home from "./ApprovalNotes/Home/Home";
import NewNFA from "./ApprovalNotes/NewNFA/NewNFA";
import Inbox from "./ApprovalNotes/Inbox/Inbox";
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
      <div className="flex flex-col" style={{ minHeight: '100vh', background: '#f1f5f9' }}>

        <Header context={context} />

        <div className="flex flex-1 overflow-hidden" style={{ marginTop: '12px', padding: '0 12px', gap: '12px' }}>

          <Sidebar />

          <main className="flex-1 overflow-auto bg-white rounded-xl shadow-sm" style={{ padding: '20px', minHeight: 0 }}>
            <Routes>
              <Route path="/"                    element={<Home context={context} sp={sp}/>} />
              <Route path="/new-nfa"             element={<NewNFA context={context} sp={sp}/>} />
              <Route path="/new-nfa/:id"         element={<NewNFA context={context} sp={sp}/>} />
              <Route path="/inbox"               element={<Inbox />} />
              <Route path="/approver-detail"     element={<ApprovalDetailsView />} />
              <Route path="/approved-rejected"   element={<ApprovedRejectedByMe />} />
              <Route path="/detail-view"         element={<DetailView />} />
              <Route path="/masters"             element={<Masters />} />
              <Route path="/function-wise-status"    element={<FunctionWiseStatus />} />
              <Route path="/approver-wise-pendency"  element={<ApproverWisePendency />} />
              <Route path="/ageing-report"       element={<AgeingReport />} />
              <Route path="/day-wise-ageing"     element={<DayWiseAgeing />} />
              <Route path="/generic-report"      element={<GenericReport />} />
            </Routes>
          </main>

        </div>
      </div>
    </HashRouter>
  );
};

export default Routing;
