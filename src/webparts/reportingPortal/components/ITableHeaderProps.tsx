export interface ITableHeaderProps {
  SearchTextChanged: (event?: React.ChangeEvent<HTMLInputElement>, newValue?: string) => void;
  DownloadSelected: (event: React.MouseEvent<HTMLButtonElement>) => void;
  SelectedRecords: any[];
  DownloadAll: (event: React.MouseEvent<HTMLButtonElement>) => void;
  SearchBoxInputText?: string;
  totalCountDisplayMsg:string;
}
