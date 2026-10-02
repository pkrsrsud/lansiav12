export type EventRow = {
  id: string;
  name: string;
  event_date: string;
  event_time: string;
  location: string;
  open_at: string;
  close_at: string;
  is_active: boolean;
  created_at: string;
  updated_at: string;
};

export type Participant = {
  id: string;
  full_name: string;
  address: string;
  is_active: boolean;
  created_at: string;
  updated_at: string;
};

export type Attendance = {
  id: string;
  event_id: string;
  participant_id: string;
  created_at: string;
  display_order: number;
  participant?: Pick<Participant, "id" | "full_name" | "address"> | null;
};

export type AdminData = {
  events: EventRow[];
  participants: Participant[];
  attendance: Attendance[];
};
