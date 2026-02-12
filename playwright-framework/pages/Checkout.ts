import { Page, Locator, Expect } from '@playwright/test';

export class Checkout {
    checkOutSteps: Locator;
    successText: Locator;
    logoutLink: Locator;
    constructor(private page: Page, private expect: Expect) {
        this.page = page;
        this.checkOutSteps = page.locator('#checkout-steps>li');
        this.successText = page.locator('div.title>strong');
        this.logoutLink = page.getByRole('link', { name: 'Log out' });
        //this.activeContinueButton = page.getByRole('button', { name: 'Continue' });
        //this.confirmButton = page.locator('#checkout-steps li.active').getByRole('button', { name: /Confirm/i });
    }

    async stepsCheckout() {
        // Wait for checkout page to load (URL and steps in DOM)
        await this.page.waitForURL(/checkout/i, { timeout: 15000 });
        await this.checkOutSteps.first().waitFor({ state: 'visible', timeout: 15000 });
        const stepCount = await this.checkOutSteps.count();
        console.log(stepCount);
       
        for(let i = 0; i < stepCount - 1; i++) {
           
            await this.checkOutSteps.nth(i).getByRole('button', { name: 'Continue' }).click();
            
        }
        //await this.confirmButton.click();
        await this.checkOutSteps.last().getByRole('button', { name: 'Confirm' }).click();
        await this.expect(this.successText).toContainText('Your order has been successfully processed!');
        await this.logoutLink.click();

    }
}