import * as React from 'react';
import { useState, useEffect } from 'react';
import { TextField, DetailsList, IColumn } from 'office-ui-fabric-react';
import { IPeopleDirectoryProps } from './IPeopleDirectoryProps';
import { IUser } from '../models/IUser';
import { Log } from '@microsoft/sp-core-library'

const PeopleDirectory: React.FC<IPeopleDirectoryProps> = ({ graphService }) => {
  const [people, setPeople] = useState<IUser[]>([]);
  const [filteredPeople, setFilteredPeople] = useState<IUser[]>([]);
  const [searchText, setSearchText] = useState("");

  useEffect(() => {
    async function fetchPeopleData() {
      Log.info('PeopleDirectory TSX','Fetching people data')
      const users = await graphService.getUsersWithPresence();
      Log.info('PeopleDirectory TSX','Fetch completed!')
      setPeople(users);
      setFilteredPeople(users);
    }

    fetchPeopleData();
  }, [graphService]);

  const onSearchChange = (event: any, text: string) => {
    setSearchText(text);
    setFilteredPeople(
      people.filter(person =>
        person.displayName.toLowerCase().includes(text.toLowerCase())
      )
    );
  };

  const columns: IColumn[] = [
    { key: 'displayName', name: 'Name', fieldName: 'displayName', minWidth: 100 },
    { key: 'jobTitle', name: 'Job Title', fieldName: 'jobTitle', minWidth: 100 },
    { key: 'department', name: 'Department', fieldName: 'department', minWidth:100},
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
      <DetailsList
        items={filteredPeople}
        columns={columns}
      />
    </div>
  );
};

export default PeopleDirectory;
