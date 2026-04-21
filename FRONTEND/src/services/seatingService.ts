import { Student, Room, TimetableEntry, SeatingPlan, Subject } from "../types";

/**
 * Seating Algorithm Rules:
 * 1. Row-wise semester separation: Each row contains students from only ONE semester.
 * 2. Same semester in one row: All students in a row belong to the same semester.
 * 3. Rows should alternate between semesters (e.g., Row 1: Sem 1, Row 2: Sem 3, Row 3: Sem 1, Row 4: Sem 5).
 * 4. AI Based Randomization: Shuffle students inside each semester group.
 */
export function generateSeatingArrangement(
  students: Student[],
  rooms: Room[],
  timetable: TimetableEntry[],
  subjects: Subject[],
  date: string,
  timeSlot: string
): SeatingPlan[] {
  // 1. Filter students who have an exam on this date and time slot
  const activeTimetable = timetable.filter(t => t.date === date && t.timeSlot === timeSlot);
  
  // Map of subjectId to Subject object for quick access
  const subjectMap = new Map<string, Subject>();
  subjects.forEach(s => subjectMap.set(s.id, s));

  // Filter students who have a subject in this shift
  const studentToSubjectMap = new Map<string, string>();
  const eligibleStudents = students.filter(st => {
    const entry = activeTimetable.find(t => {
      const sub = subjectMap.get(t.subjectId);
      if (!sub) return false;
      
      return st.programmeId === sub.programmeId && 
             st.semester === sub.semester &&
             (sub.specialization === "General" || !sub.specialization || st.specialization === sub.specialization);
    });
    
    if (entry) {
      studentToSubjectMap.set(st.id, entry.subjectId);
      return true;
    }
    return false;
  });

  // 2. Group students by semester
  const semesterGroups: Record<number, Student[]> = {};
  eligibleStudents.forEach(s => {
    if (!semesterGroups[s.semester]) semesterGroups[s.semester] = [];
    semesterGroups[s.semester].push(s);
  });

  // 3. Group by specialization and then randomize within each group to ensure they sit together
  Object.keys(semesterGroups).forEach(sem => {
    const s = Number(sem);
    const studentsInSem = semesterGroups[s];
    
    // Group by specialization
    const specGroups: Record<string, Student[]> = {};
    studentsInSem.forEach(st => {
      const spec = (st.specialization || "General").trim();
      if (!specGroups[spec]) specGroups[spec] = [];
      specGroups[spec].push(st);
    });
    
    // Shuffle each specialization group and combine
    let processedStudents: Student[] = [];
    // Sort specialization names to have consistent ordering
    Object.keys(specGroups).sort().forEach(spec => {
      processedStudents = [...processedStudents, ...shuffleArray(specGroups[spec])];
    });
    
    semesterGroups[s] = processedStudents;
  });

  const availableSemesters = Object.keys(semesterGroups).map(Number).sort((a, b) => a - b);
  if (availableSemesters.length === 0) return [];

  const seatingPlans: SeatingPlan[] = [];
  let currentRoomIndex = 0;
  let currentSemesterIndex = 0;

  // 4. Assign students to rooms column by column
  while (currentRoomIndex < rooms.length && availableSemesters.some(sem => semesterGroups[sem].length > 0)) {
    const room = rooms[currentRoomIndex];
    
    // Safety check for old room data
    if (!room.columns || !room.seatsPerColumn) {
      currentRoomIndex++;
      continue;
    }
    
    // Iterate through columns
    for (let col = 1; col <= room.columns; col++) {
      const maxSeatsInCol = room.seatsPerColumn[col - 1];
      if (!maxSeatsInCol) continue;
      
      // Pick a semester for this column (alternating)
      let attempts = 0;
      while (semesterGroups[availableSemesters[currentSemesterIndex]].length === 0 && attempts < availableSemesters.length) {
        currentSemesterIndex = (currentSemesterIndex + 1) % availableSemesters.length;
        attempts++;
      }

      const currentSem = availableSemesters[currentSemesterIndex];
      const studentsInSem = semesterGroups[currentSem];

      if (studentsInSem.length > 0) {
        for (let s = 1; s <= maxSeatsInCol; s++) {
          if (studentsInSem.length > 0) {
            const student = studentsInSem.shift()!;
            const subjectId = studentToSubjectMap.get(student.id)!;
            
            seatingPlans.push({
              id: `${date}-${timeSlot}-${room.id}-${col}-${s}`,
              date,
              timeSlot,
              roomId: room.id,
              row: s, // Using 'row' as the seat index within the column
              seat: col, // Using 'seat' as the column index
              studentId: student.id,
              subjectId
            });
          }
        }
      }

      // Move to next semester for the next column
      currentSemesterIndex = (currentSemesterIndex + 1) % availableSemesters.length;
    }

    currentRoomIndex++;
  }

  return seatingPlans;
}

function shuffleArray<T>(array: T[]): T[] {
  const newArray = [...array];
  for (let i = newArray.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [newArray[i], newArray[j]] = [newArray[j], newArray[i]];
  }
  return newArray;
}
