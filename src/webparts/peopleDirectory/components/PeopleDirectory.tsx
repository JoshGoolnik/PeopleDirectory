import * as React from 'react';
import { useState, useEffect } from 'react';
import { TextField, DetailsList, IColumn } from 'office-ui-fabric-react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IPeopleDirectoryProps } from './IPeopleDirectoryProps';
import { IUser } from '../models/IUser';

// Define SVGs for status
const statusIcons = {
  available: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="13" viewBox="0 0 10 10" fill="#92c353">
                <circle cx="5" cy="5" r="5"/>
                <polyline points="2,4 4,7 8,3" style="fill:none;stroke:white;stroke-width:1" />
              </svg>`,
  away: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="13" viewBox="0 0 10 10" fill="#fcd116">
          <circle cx="5" cy="5" r="5"/>
          <polyline points="5,3 5,6 7,7 " style="fill:none;stroke:white;stroke-width:1"/>
        </svg>`,
  busy: '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="13" viewBox="0 0 10 10" fill="#c4314b"><circle cx="5" cy="5" r="5"/></svg>',
  dnd: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="13" viewBox="0 0 10 10" fill="#c4314b">
          <circle cx="5" cy="5" r="5"/>
          <line x1="1" y1="5" x2="9" y2="5" style="stroke:white;stroke-width:2" />
        </svg>`,
  ooo: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="13" viewBox="0 0 10 10" stroke="#b4009e" fill="none">
          <circle cx="5" cy="5" r="5"/>
          <polyline points="5,2 2,5 5,8"/>
          <line x1 = "2" y1="5" x2="8"  y2="5">
        </svg>`,
  offline: `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="13" viewBox="0 0 10 10" stroke="#959595" fill="none">
              <circle cx="5" cy="5" r="5"/>
              <line x1 = "2" y1="2" x2="8"  y2="8" />
              <line x1 = "2" y1="8" x2="8"  y2="2" />
            </svg>`
};


const PeopleDirectory: React.FC<IPeopleDirectoryProps> = ({ graphService }) => {
  const [people, setPeople] = useState<IUser[]>([]);
  const [filteredPeople, setFilteredPeople] = useState<IUser[]>([]);
  const [searchText, setSearchText] = useState("");
  const [departments, setDepartments] = useState<IDropdownOption[]>([]);
  const [selectedDepartment, setSelectedDepartment] = useState<string | undefined>(undefined);

  useEffect(() => {
    async function fetchPeopleData() {
      const users = await graphService.getUsersWithPresence();

      // Extract unique departments
      const uniqueDepartments = Array.from(new Set(users.map(user => user.department)))
        .filter(department => department) // Filter out undefined departments
        .map(department => ({ key: department, text: department }));

      setPeople(users);
      setFilteredPeople(users);
      setDepartments(uniqueDepartments);
    }

    fetchPeopleData();
  }, [graphService]);

  const onSearchChange = (event: any, text: string) => {
    setSearchText(text);
    filterPeople(text, selectedDepartment);
  };

  const onDepartmentChange = (event: any, option?: IDropdownOption) => {
    setSelectedDepartment(option?.key as string);
    filterPeople(searchText, option?.key as string);
  };

  const filterPeople = (search: string, department?: string) => {
    const filtered = people.filter(person => 
      person.displayName.toLowerCase().includes(search.toLowerCase()) &&
      (!department || person.department === department) // Filter by department if selected
    );
    setFilteredPeople(filtered);
  };

 // Function to render the custom SVG status icon
 const renderStatusIcon = (availability: string) => {
  switch (availability.toLowerCase()) {
    case 'available':
      return <span dangerouslySetInnerHTML={{ __html: statusIcons.available }} />;
    case 'away':
      return <span dangerouslySetInnerHTML={{ __html: statusIcons.away }} />;
    case 'busy':
      return <span dangerouslySetInnerHTML={{ __html: statusIcons.busy }} />;
    case 'donotdisturb':
      return <span dangerouslySetInnerHTML={{ __html: statusIcons.dnd }} />;
    case 'outofoffice':
      return <span dangerouslySetInnerHTML={{ __html: statusIcons.ooo }} />;
    case 'offline':
        return <span dangerouslySetInnerHTML={{ __html: statusIcons.offline }} />;
    default:
      return availability;
  }
};

// Function to remove HTML tags, classes and css properties.
const regex = /(<([^>]+)>)/gi;
const formatStatusMessage = (statusMessage: string) => {
  return statusMessage.replace(regex,'');
}
  const columns: IColumn[] = [
    { key: 'displayName', name: 'Name', fieldName: 'displayName', minWidth: 120, maxWidth: 200, isResizable: true},
    { key: 'jobTitle', name: 'Job Title', fieldName: 'jobTitle', minWidth: 75, maxWidth: 220, isResizable: true},
    { key: 'department', name: 'Department', fieldName: 'department', minWidth: 75, maxWidth: 220, isResizable: true},
    {
      key: 'availability',
      name: '?',
      fieldName: 'availability',
      minWidth:30, 
      maxWidth:30,
      onRender: (item: IUser) => renderStatusIcon(item.availability)
    },
    { key: 'activity', name: 'Activity', fieldName: 'activity', minWidth: 60, maxWidth: 160, isResizable: true},
    { key: 'statusMessage', name: 'Status Message', fieldName: 'statusMessage', minWidth: 220, maxWidth: 500, isResizable: true, onRender: (item:IUser) => formatStatusMessage(item.statusMessage || '')}
  ];

  return (
    <div>
      <TextField
        label="Search by name"
        value={searchText}
        onChange={onSearchChange}
      />
      <Dropdown
        placeholder="Select a department"
        label="Filter by Department"
        options={[{ key: '', text: 'All Departments' }, ...departments]}
        onChange={onDepartmentChange}
      />
      <DetailsList
        items={filteredPeople}
        columns={columns}
      />
    </div>
  );
};

export default PeopleDirectory;
