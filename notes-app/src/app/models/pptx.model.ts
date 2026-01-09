export type PptxObjectBase = {
  id: string;
  type: 'text' | 'image';
  x: number;  // inches
  y: number;  // inches
  w: number;  // inches
  h: number;  // inches
  z: number;
  rotation?: number;
};

export type PptxTextObject = PptxObjectBase & {
  type: 'text';
  text: string;
  fontSize?: number;
};

export type PptxImageObject = PptxObjectBase & {
  type: 'image';
  src: string; // URL (served by notes-server)
};

export type PptxObject = PptxTextObject | PptxImageObject;

export type PptxSlide = {
  index: number;
  objects: PptxObject[];
};

export type ImportPptxResponse = {
  ok: boolean;
  deckId: string;
  slides: PptxSlide[];
};
