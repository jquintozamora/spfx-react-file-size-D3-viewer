export interface ITreeMapNode {
  name: string;
  value?: number;
  id?: string;
  children?: ITreeMapNode[];
}
