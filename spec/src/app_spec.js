import ZAFClient from 'zendesk_app_framework_sdk';
import App from '../../src/javascripts/app';

describe('App', () => {
  let app;

  beforeEach(() => {
    let client = ZAFClient.init();
    app = new App(client, { metadata: {}, context: {} });
  });

  describe('#renderMain', () => {
    beforeEach(() => {
      spyOn(app, 'switchTo');
    });

    it('switches to the main template', () => {
      var data = { user: 'Mikkel' };
      app.renderMain(data);
      expect(app.switchTo).toHaveBeenCalledWith('main', data.user);
    });
  });
});
