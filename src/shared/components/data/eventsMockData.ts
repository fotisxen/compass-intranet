export interface IEvent {
  date: string;
  title: string;
  initials: string;
  imageUrl?: string;
}

export const events: IEvent[] = [
  {
    date: '14.09.2026',
    title: 'Marketing Workshop',
    initials: 'MW',
    imageUrl: 'https://images.unsplash.com/photo-1552664730-d307ca884978?w=400&q=80'
  }
];
