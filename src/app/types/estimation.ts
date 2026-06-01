export interface Material {
  name: string;
  unitPrice: number;
  quantity: number;
}

export interface LaborTask {
  description: string;
  hours: number;
  hourlyRate: number;
}

export interface CalculationInput {
  materials: Material[];
  labor: LaborTask[];
  marginPercentage: number;
  customerName?: string;
  metrics?: {
    peakKW?: number;
    dayConsumptionKWh?: number;
    nightConsumptionKWh?: number;
    apartmentCount?: number;
    totalDeviceCount?: number;
  };
}