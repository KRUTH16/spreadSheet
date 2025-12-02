export type CellStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  bgColor?: string;
  color?:string;
  align?: 'left' | 'center' | 'right';
};

export type Cell = {
  value: string;      
  formula?: string;  
  style?: CellStyle;
};

export type Sheet = Cell[][];
