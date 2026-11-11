import * as React from "react";
import { Routes, Route, HashRouter, Navigate, useLocation } from "react-router-dom";
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
import ApprovedForFinanceView from "./ApprovalNotes/ApprovedForFinance/ApprovedForFinanceView"
import DetailView from "./AdminSection/DetailView/DetailView";
import Masters from "./AdminSection/Masters/Masters";
import FunctionWiseStatus from "./Reports/FunctionWiseStatus/FunctionWiseStatus";
import ApproverWisePendency from "./Reports/ApproverWisePendency/ApproverWisePendency";
import AgeingReport from "./Reports/AgeingReport/AgeingReport";
import DayWiseAgeing from "./Reports/DayWiseAgeing/DayWiseAgeing";
import GenericReport from "./Reports/GenericReport/GenericReport";

import { forceScrollToTop } from "../../../Share/utils";

const ScrollToTop: React.FC = () => {
  const { pathname } = useLocation();

  React.useEffect(() => {
    // Run immediately
    forceScrollToTop();

    // Periodically run for 1.5s after a path change to battle layout updates and browser scroll preservation
    let count = 0;
    const interval = setInterval(() => {
      forceScrollToTop();
      count++;
      if (count >= 15) {
        clearInterval(interval);
      }
    }, 100);

    return () => {
      clearInterval(interval);
    };
  }, [pathname]);

  return null;
};

const Routing: React.FC<INfaProps> = ({ context, sp }) => {
  React.useEffect(() => {
    document.body.classList.add("nfa-portal-active");
    return () => {
      document.body.classList.remove("nfa-portal-active");
    };
  }, []);

  return (
    <HashRouter>
      <ScrollToTop />
      <div className="w-full flex flex-col min-h-screen" style={{ background: 'var(--page-background, #FAF7F0)' }}>
        <Header context={context} />

        <div className="w-full flex flex-1 min-w-0" style={{ padding: '10px' }}>
          <Sidebar context={context} sp={sp} />

          <main className="flex-1 w-full min-w-0" style={{ padding: '0 0 30px 20px' }}>
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
              <Route path="/approved-nfa-finance" element={<ApprovedForFinanceView context={context} sp={sp}/>} />
              <Route path="/detail-view"         element={<DetailView context={context} sp={sp}/>} />
              <Route path="/masters"             element={<Masters />} />
              <Route path="/function-wise-status"    element={<FunctionWiseStatus context={context} sp={sp}/>} />
              <Route path="/approver-wise-pendency"  element={<ApproverWisePendency context={context} sp={sp}/>} />
              <Route path="/ageing-report"       element={<AgeingReport context={context} sp={sp}/>} />
              <Route path="/day-wise-ageing"     element={<DayWiseAgeing context={context} sp={sp}/>} />
              <Route path="/generic-report"      element={<GenericReport context={context} sp={sp}/>} />
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
