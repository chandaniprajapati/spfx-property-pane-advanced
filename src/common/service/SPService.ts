import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp/presets/all";

export class SPService {
    constructor(private context: WebPartContext) {
        sp.setup({
            spfxContext: this.context
        });
    }

    public async getFields(selectedList: string): Promise<any> {
        try {
            const allFields: any[] = await sp.web.lists
                .getById(selectedList)
                .fields
                .filter("Hidden eq false and ReadOnlyField eq false")
                .get()
            return allFields;
        }
        catch (err) {
            Promise.reject(err);
        }
    }

    public async getItems(): Promise<any> {
        try {
            const items: any[] = await sp.web.lists
                .getByTitle('DayInfo')
                .items
                .get()
            return items;
        }
        catch (err) {
            Promise.reject(err);
        }
    }
}