export type ContrastSortMode = "createdAtDesc" | "updatedAtDesc";

export type ContrastRow = {
  id: number;
  standardSamplePath: string;
  samplePath: string;
  analysisResultsPath: string;
  thresholdNumber: string;
  lastResultFilePath: string;
  remarks: string;
  createdAt: number;
  updatedAt: number;
};
