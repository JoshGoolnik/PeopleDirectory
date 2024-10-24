import * as React from 'react';
import { useState, useEffect } from 'react';
import { TextField, DetailsList, IColumn } from 'office-ui-fabric-react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IPeopleDirectoryProps } from './IPeopleDirectoryProps';
import { IUser } from '../models/IUser';
import { AvailableIcon, AwayIcon, BusyIcon, DoNotDisturbIcon, OutOfOfficeIcon, OfflineIcon } from './StatusIcons';

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
      return <AvailableIcon />;
    case 'availableidle':
    case 'away':
    case 'berightback':
      return <AwayIcon />;
    case 'busy':
      return <BusyIcon/>;
    case 'donotdisturb':
    case 'dnd':
    case 'do not disturb':
      return <DoNotDisturbIcon />;
    case 'outofoffice':
    case 'out of office':
      return <OutOfOfficeIcon />;
    case 'offline':
        return <OfflineIcon />;
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
    { key: 'displayName', name: 'Name', fieldName: 'displayName', minWidth: 120, maxWidth: 140, isResizable: true, isMultiline:true},
    { key: 'jobTitle', name: 'Job Title', fieldName: 'jobTitle', minWidth: 75, maxWidth: 200, isResizable: true, isMultiline:true},
    { key: 'department', name: 'Department', fieldName: 'department', minWidth: 75, maxWidth: 160, isResizable: true, isMultiline:true},
    {
      key: 'availability',
      name: '?',
      fieldName: 'availability',
      minWidth:30, 
      maxWidth:30,
      onRender: (item: IUser) => renderStatusIcon(item.availability)
    },
    { key: 'activity', name: 'Activity', fieldName: 'activity', minWidth: 60, maxWidth: 110, isResizable: true, isMultiline:true},
    { key: 'statusMessage', name: 'Status Message', fieldName: 'statusMessage', minWidth: 220, maxWidth: 520, isResizable: true, isMultiline:true, onRender: (item:IUser) => formatStatusMessage(item.statusMessage || '')}
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
