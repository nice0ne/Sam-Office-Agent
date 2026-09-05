import { HostType } from '../../types';
import { ExcelDriver } from './excelDriver';
import { MockOfficeDriver } from './mockDriver';
import { PPTDriver } from './pptDriver';
import { IDocumentDriver } from './types';
import { WordDriver } from './wordDriver';

let activeDriver: IDocumentDriver | null = null;

export function getOfficeDriver(_host?: HostType): IDocumentDriver {
  if (activeDriver) return activeDriver;

  const office = typeof window !== 'undefined' ? (window as any).Office : (globalThis as any).Office;

  if (typeof office !== 'undefined' && office?.context?.host) {
    const officeHost = office.context.host;
    if (officeHost === office.HostType?.Excel || officeHost === 'Excel') {
      activeDriver = new ExcelDriver() as unknown as IDocumentDriver;
    } else if (officeHost === office.HostType?.Word || officeHost === 'Word') {
      activeDriver = new WordDriver() as unknown as IDocumentDriver;
    } else if (officeHost === office.HostType?.PowerPoint || officeHost === 'PowerPoint') {
      activeDriver = new PPTDriver() as unknown as IDocumentDriver;
    }
  }

  if (!activeDriver) {
    activeDriver = new MockOfficeDriver();
  }

  return activeDriver;
}

export function resetOfficeDriver(): void {
  activeDriver = null;
}

export function setOfficeDriver(driver: IDocumentDriver | null): void {
  activeDriver = driver;
}

export * from './types';
export * from './mockDriver';
export * from './excelDriver';
export * from './wordDriver';
export * from './pptDriver';
