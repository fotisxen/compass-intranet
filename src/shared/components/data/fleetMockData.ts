export interface IOffice {
  city: string;
  country: string;
  lat: number;
  lng: number;
  address: string;
}

export const offices: IOffice[] = [
  { city: 'Athens', country: 'Greece', lat: 37.9838, lng: 23.7275, address: 'HQ — Vasilissis Sofias Ave 1' },
  { city: 'London', country: 'United Kingdom', lat: 51.5072, lng: -0.1276, address: '10 Canada Square' },
  { city: 'New York', country: 'United States', lat: 40.7128, lng: -74.006, address: '350 5th Ave' },
  { city: 'Singapore', country: 'Singapore', lat: 1.3521, lng: 103.8198, address: '1 Raffles Place' },
  { city: 'Sydney', country: 'Australia', lat: -33.8688, lng: 151.2093, address: '1 Macquarie Pl' }
];
