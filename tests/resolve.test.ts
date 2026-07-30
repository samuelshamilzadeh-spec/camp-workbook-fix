import { describe, expect, it } from 'vitest';
import { parsePersonalUrl } from '../scripts/resolve-workbook';

describe('parsePersonalUrl', () => {
  it('recovers the owner UPN and file name from a OneDrive browser URL', () => {
    const url =
      'https://comprehensivecommunity-my.sharepoint.com/:x:/r/personal/tplaut_premierassist_com/' +
      '_layouts/15/doc2.aspx?sourcedoc=%7B27BDB9A2-4900-4841-8B1A-0865FEE7EC90%7D' +
      '&file=CAMP%20PATIENT%20SIGN%20UP%20SHEET.xlsx&action=default&mobileredirect=true';

    expect(parsePersonalUrl(url)).toEqual({
      upn: 'tplaut@premierassist.com',
      fileName: 'CAMP PATIENT SIGN UP SHEET.xlsx',
    });
  });

  it('handles a multi-label domain', () => {
    expect(parsePersonalUrl('https://x-my.sharepoint.com/personal/jo_smith_co_uk/Documents/')).toEqual(
      { upn: 'jo@smith.co.uk', fileName: undefined },
    );
  });

  it('returns undefined for a normal SharePoint site URL', () => {
    expect(
      parsePersonalUrl('https://contoso.sharepoint.com/sites/CampOps/Shared%20Documents/Book.xlsx'),
    ).toBeUndefined();
  });
});
