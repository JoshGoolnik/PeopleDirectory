import * as React from 'react';
import { useState, useEffect } from 'react';
import { TextField, DetailsList, IColumn } from 'office-ui-fabric-react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IPeopleDirectoryProps } from './IPeopleDirectoryProps';
import { IUser } from '../models/IUser';
import { Log } from '@microsoft/sp-core-library';

const PeopleDirectory: React.FC<IPeopleDirectoryProps> = ({ graphService }) => {
  const [people, setPeople] = useState<IUser[]>([]);
  const [filteredPeople, setFilteredPeople] = useState<IUser[]>([]);
  const [searchText, setSearchText] = useState("");
  const [departments, setDepartments] = useState<IDropdownOption[]>([]);
  const [selectedDepartment, setSelectedDepartment] = useState<string | undefined>(undefined);

  useEffect(() => {
    async function fetchPeopleData() {
      Log.info('PeopleDirectory TSX', 'Fetching people data');
      const users = await graphService.getUsersWithPresence();
      Log.info('PeopleDirectory TSX', 'Fetch completed!');

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

  const columns: IColumn[] = [
    { key: 'displayName', name: 'Name', fieldName: 'displayName', minWidth: 100 },
    { key: 'jobTitle', name: 'Job Title', fieldName: 'jobTitle', minWidth: 100 },
    { key: 'department', name: 'Department', fieldName: 'department', minWidth: 100 },
    { key: 'availability', name: 'Availability', fieldName: 'availability', minWidth: 100 },
    { key: 'activity', name: 'Activity', fieldName: 'activity', minWidth: 100 },
    { key: 'statusMessage', name: 'Status Message', fieldName: 'statusMessage', minWidth: 200 },
    { key: 'workLocation', name: 'Work Location', fieldName: 'workLocation', minWidth: 100 }
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
