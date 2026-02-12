import { test, expect } from '@playwright/test';
import { POManager } from '../POManager/POManager';
import data from '../JsonData/data.json';

test('Demowebshop Login', async ({ page }) => {
  const pomManager = new POManager(page, expect);
  await pomManager.loginPage.siteURL(data.URL);
  await pomManager.loginPage.login(data.Username, data.Password);
  await pomManager.loginPage.search(data.Search, data.Product);
  await pomManager.cart.goToCart();
  await pomManager.cart.checkCart(data.Product);
  await pomManager.cart.checkCartQuantity();
  await pomManager.checkout.stepsCheckout();
  
 
});