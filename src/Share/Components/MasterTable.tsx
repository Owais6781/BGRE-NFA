import * as React from 'react';
import { useState } from 'react';
import { ChevronUp, ChevronDown, ExternalLink } from 'lucide-react';
import TableToolbar from './TableToolbar';
import { SearchInput } from './SearchInput';
import PageTitle from './PageTitle';
import "../styles/global.scss";

interface MasterItem { name: string; url: string; }

interface MasterTableProps {
  title:      string;
  data:       MasterItem[];
  nameLabel?: string;
  urlLabel?:  string;
  nameWidth?: string;
}

const MasterTable: React.FC<MasterTableProps> = ({
  title, data, nameLabel = 'Master', urlLabel = 'URL', nameWidth = '40%',
}) => {
  const [search,      setSearch]      = useState('');
  const [isCollapsed, setIsCollapsed] = useState(false);

  const filtered = data.filter((item) =>
    item.name.toLowerCase().includes(search.toLowerCase()) ||
    item.url.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div className="pageContainer">
      <div className="contentCard">

        <PageTitle 
          title={title} 
          subtitle="Manage and configure system masters and mappings." 
          action={
            <button className="iconBtn" onClick={() => setIsCollapsed(c => !c)} title={isCollapsed ? 'Expand' : 'Collapse'}>
              {isCollapsed ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
            </button>
          }
        />

        {!isCollapsed && (
          <>
            <TableToolbar
              left={<SearchInput value={search} onChange={setSearch} placeholder="Search masters…" />}
            />

            <div className="tableWrapper">
              <table className="commonTable">
                <thead>
                  <tr>
                    <th className="tableHeader" style={{ width: nameWidth }}>{nameLabel}</th>
                    <th className="tableHeader">{urlLabel}</th>
                  </tr>
                </thead>
                <tbody>
                  {filtered.length === 0 ? (
                    <tr>
                      <td colSpan={2} className="tableCell text-center py-8" style={{ color: '#94a3b8' }}>
                        No records found
                      </td>
                    </tr>
                  ) : (
                    filtered.map((item, i) => (
                      <tr key={i} className="tableRow">
                        <td className="tableCell">
                          <a href={item.url} target="_blank" rel="noreferrer" className="mastersLink flex items-center gap-1">
                            {item.name}
                            <ExternalLink size={11} />
                          </a>
                        </td>
                        <td className="tableCell" style={{ color: '#64748b', fontSize: 11 }}>
                          {item.url}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </>
        )}

      </div>
    </div>
  );
};

export default MasterTable;
