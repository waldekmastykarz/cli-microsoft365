import { Hash, Occurrence } from "./";

export interface Finding {
  description: string;
  id: string;
  occurrences: Occurrence[];
  properties: Hash;
  resolutionType: string;
  severity: string;
  supersedes: string[];
  title: string;
}