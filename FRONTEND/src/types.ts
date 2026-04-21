export interface Programme {
  id: string;
  name: string;
}

export interface Subject {
  id: string;
  programmeId: string;
  semester: number;
  code: string;
  name: string;
  type: "Theory" | "Practical";
  specialization: string;
}

export interface Student {
  id: string;
  studentId: string;
  name: string;
  programmeId: string;
  specialization: string;
  semester: number;
}

export interface Room {
  id: string;
  roomNumber: string;
  capacity: number;
  columns: number;
  seatsPerColumn: number[];
}

export interface TimetableEntry {
  id: string;
  date: string;
  timeSlot: string;
  programmeId: string;
  semester: number;
  subjectId: string;
}

export interface SeatingPlan {
  id: string;
  date: string;
  timeSlot: string;
  roomId: string;
  row: number;
  seat: number;
  studentId: string;
  subjectId: string;
}

