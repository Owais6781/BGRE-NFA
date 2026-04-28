import * as React          from 'react';
import MasterTable         from '../../../../../Share/Components/MasterTable';
import "../../../../../Share/styles/global.scss";

const MASTERS_DATA = [
  { name: 'Approver Master',        url: '/sites/NFA-CORP/SitePages/ApproverMaster.aspx'          },
  { name: 'Entity Master',          url: '/sites/NFA-CORP/SitePages/EntityMaster.aspx'            },
  { name: 'Function Master',        url: '/sites/NFA-CORP/SitePages/FunctionMaster.aspx'          },
  { name: 'Project Master',         url: '/sites/NFA-CORP/SitePages/Projectmaster.aspx'           },
  { name: 'Project Entity Mapping', url: '/sites/NFA-CORP/SitePages/ProjectSubEntityMapping.aspx' },
  { name: 'DOA Master',             url: '/sites/NFA-CORP/sitepages/DOAMaster.aspx'               },
  { name: 'Head/Sub-Head',          url: '/sites/NFA-CORP/sitepages/HeadSubMaster.aspx'           },
];

const Masters: React.FC = () => {
  return (
    <MasterTable
      title="All Masters"
      data={MASTERS_DATA}
      nameLabel="Master"
      urlLabel="URL"
      nameWidth="40%"
    />
  );
};

export default Masters;