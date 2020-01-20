const { Builder, By, until, error } = require('selenium-webdriver');
const { Options } = require('selenium-webdriver/chrome');
const XLSX = require('xlsx');

const { TimeoutError } = error;

const locations = {
  meetingRoom1: 'c7afa48e-7874-447f-8e04-c650ff7926ee',
  meetingRoom10: '98a2ab63-2976-47ca-95f5-a4499bd70e52',
  meetingRoom2: 'e1a6c40f-30dd-4c55-9752-9180bc6abbda',
  meetingRoom3: '09a2992d-5203-40e7-aa01-be892bfcbb4b',
  meetingRoom4: 'c1d6a642-71d4-494f-8f9a-735a3776c70b',
  meetingRoom5: 'f86503ff-4547-437b-bedd-d3522995e967',
  meetingRoom6: '678fdd7c-7af4-4bd8-8ea5-1a6d9a74f162'
};

const slots = {
  meetingRoom1:
    '%2fX7xxd0fq5A9T5D1PEcpXEbMTrKQ5VR1ajpbY13KVe8TCI9wv7T425X7AkjEbmLThwTkBHDAtBzy2MR%2boLpzszt9lQuswXWAwAXU%2bFAuRwE%3d',
  meetingRoom10:
    '3hJXiZ5GuIsAHkTS88LvFwWDEEYskVFHlLFWlxCTnGwc%2blxSqKkLBU5y1L7zcCHZ7r4snNGTOqkZEVZDLf%2bT0kS2y%2bJzQejCTNcVYgBGORs%3d',
  meetingRoom2:
    'aRl0FmeDfBJqsFSRBi0ncw%2bPQx%2baq8UtdUnwwGXg7%2bfMSl2bwP5ui2mCAdCixYXiKhjPC3%2f3S7O9%2fedlK9SXROpnVsoG6SEWpuLIhm5cDW0%3d',
  meetingRoom3:
    'b0uWqQ%2f2ae%2fmzk2NidV1LpJ%2b%2fLZ67%2bhNqjM6dG%2fvDDWNifDH2szD%2fA7tm0ERv6hqofpVvg%2fXxqhFMN4NXt7mi0%2bk9uRW3exGxK9PW7Z3S%2fE%3d',
  meetingRoom4:
    'RMp6cf8GL%2baYsg%2foC28slHiLXjq4WrzGyIMXOWlHz0HcOD0PynC3wf%2bZyPG8UThrUq0A6oUKyzqNuhSyH5iE5PwRTrKWfqkbhQ1i6PIT%2fWQ%3d',
  meetingRoom5:
    'Js1u2c883CxQ2MyUJyH%2fUZvxmR2i1ClDevCesAUCpe7ZdFFBjfmDqQ5jaMx71MzTj8Jw%2b2jT0HERO3HORAl3s%2bO8gFHNiJ%2ba7u3Z5s5qffA%3d',
  meetingRoom6:
    'Le7qQhZ66UNfdejEyBh4piCv4zRqEzhtG%2bjXRtyd2VBQSshiDfPSSz7leDNg9APQtDEGZY52DkMkuAyIESrPeX0nVfQIS9FFkTLjeNUxTWM%3d'
};

const getAuthenticatedInstance = async (username, password) => {
  const driver = await new Builder()
    .forBrowser('chrome')
    .setChromeOptions(new Options().headless())
    .build();

  await driver
    .get(`https://${username}:${password}@utownfbs.nus.edu.sg`)
    .then(() => driver.get('https://utownfbs.nus.edu.sg/utown/apptop.aspx'));

  return driver;
};

const goToAvailabilityView = async (driver, venue) => {
  const frame = await driver.wait(until.elementLocated(By.id('frameBottom')), 20000);

  await driver.switchTo().frame(frame);
  await driver.sleep(1000);

  const facilitySelect = await driver.wait(
    until.elementLocated(By.name('FacilityType$ctl02')),
    20000
  );

  //////////////////////////////////////////////////////////////

  await facilitySelect.click();
  await driver.sleep(1000);

  const facilityEmptyOption = await driver.findElement(By.css('option[value=""]'));
  await facilityEmptyOption.click();
  await driver.sleep(1000);

  const savePreference = await driver.findElement(By.linkText('Save Preference'));
  await savePreference.click();

  await driver.wait(until.alertIsPresent(), 20000);

  await driver
    .switchTo()
    .alert()
    .accept();
  await driver.sleep(1000);

  await driver.switchTo().defaultContent();
  await driver.sleep(1000);
  await driver.switchTo().frame(frame);
  await driver.sleep(1000);

  //////////////////////////////////////////////////////////////

  const facilitySelectNew = await driver.wait(
    until.elementLocated(By.name('FacilityType$ctl02')),
    20000
  );
  await facilitySelectNew.click();
  await driver.sleep(1000);

  const facilityOption = await driver.wait(
    until.elementLocated(By.css('option[value="803b9748-784e-4c0e-9f04-fdf2b78728f7"]')),
    20000
  );
  await facilityOption.click();
  await driver.sleep(1000);

  const venueOption = await driver.wait(
    until.elementLocated(By.css(`option[value="${locations[venue]}"]`)),
    20000
  );
  await venueOption.click();
  await driver.sleep(1000);

  const viewAvailability = await driver.findElement(By.id('btnViewAvailability'));
  await viewAvailability.click();
};

const bookSlots = async (driver, from, to) => {
  const fromTimeSelect = await driver.wait(until.elementLocated(By.name('from$ctl02')), 20000);
  const staleToTimeSelect = await driver.findElement(By.name('to$ctl02'));
  await fromTimeSelect.click();
  const fromTimeOption = await driver.findElement(
    By.css(`select[name="from$ctl02"] > option[value="1800/01/01 ${from}"]`)
  );
  await fromTimeOption.click();

  await driver.wait(until.stalenessOf(staleToTimeSelect), 20000);
  const newToTimeSelect = await driver.findElement(By.name('to$ctl02'));
  await newToTimeSelect.click();
  const toTimeOption = await driver.findElement(
    By.css(`select[name="to$ctl02"] > option[value="1800/01/01 ${to}"]`)
  );
  await toTimeOption.click();
  await driver.sleep(1000);
  const expectedAttendants = await driver.findElement(By.name('ExpectedNoAttendees$ctl02'));
  await expectedAttendants.clear();
  await expectedAttendants.sendKeys('6');
  await driver.sleep(1000);
  const usageTypeSelect = await driver.findElement(By.name('UsageType$ctl02'));
  await usageTypeSelect.click();
  const usageTypeOption = await driver.findElement(
    By.css('option[value="d946c992-97e3-4a44-bb11-07ad0440563d"]')
  );
  await usageTypeOption.click();
  await driver.sleep(1000);
  const chargeGroupSelect = await driver.findElement(By.name('ChargeGroup$ctl02'));
  await chargeGroupSelect.click();
  const chargeGroupOption = await driver.findElement(By.css('option[value="0"]'));
  await chargeGroupOption.click();
  await driver.sleep(1000);
  const purpose = await driver.findElement(By.name('Purpose$ctl02'));
  await purpose.clear();
  await purpose.sendKeys('Study');

  const createBooking = await driver.findElement(By.id('btnCreateBooking'));
  await createBooking.click();
};

const book = async (username, password, bookingDate, bookingTimes, venue) => {
  let driver;
  try {
    driver = await getAuthenticatedInstance(username, password);

    await goToAvailabilityView(driver, venue);

    //////////////////////////////////////////////////////////////

    const labelMessage = await driver.wait(until.elementLocated(By.id('lblMessage')), 20000);
    await driver.wait(until.elementIsVisible(labelMessage), 20000);
    await driver.sleep(1000);

    const targetSlot = await driver.wait(
      until.elementLocated(By.id(`${bookingDate}_0_${slots[venue]}`)),
      20000
    );
    await targetSlot.click();

    const frame = await driver.wait(until.elementLocated(By.id('frmCreate')), 20000);
    await driver.switchTo().frame(frame);
    await driver.sleep(1000);

    //////////////////////////////////////////////////////////////

    const results = [];
    for (const bookingTime of bookingTimes) {
      const [from, to] = bookingTime;
      await bookSlots(driver, from, to);

      try {
        await driver.wait(
          until.elementLocated(
            By.xpath("//span[text()='The specified slot is booked by another user']")
          ),
          20000
        );
        results.push(`${username}: ${venue} ${bookingTime.join(' - ')} Failed`);
      } catch (e) {
        if (e.message.includes('Waiting')) {
          results.push(`${username}: ${venue} ${bookingTime.join(' - ')} Success`);

          const createAnotherBooking = await driver.wait(
            until.elementLocated(By.id('btnCreateAnotherBooking')),
            20000
          );
          await createAnotherBooking.click();
        } else {
          console.log(username, venue, bookingTime, e);
        }
      }
    }
    driver.close();

    return results;
  } catch (e) {
    console.log(username, venue, bookingTimes);

    if (e instanceof TimeoutError) {
      console.log(e.name);
      driver && driver.close();
    } else {
      console.log(e);
    }

    return bookingTimes.map(
      bookingTime => `${username}: ${venue} ${bookingTime.join(' - ')} Failed`
    );
  }
};

const nextAvailableDay = 6 * 24 * 3600 * 1000;
const todayDate = new Date();
const targetDate = new Date(todayDate.getTime() + nextAvailableDay);

const year = targetDate.getFullYear();
const month = `0${targetDate.getMonth() + 1}`.slice(-2);
const date = `0${targetDate.getDate()}`.slice(-2);

const bookingDate = [year, month, date].join('');

const workbook = XLSX.readFile('UTown FBS.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

const endCell = XLSX.utils.decode_range(worksheet['!ref']).e;

const order = {};
for (i = 1; i < endCell.r + 1; i++) {
  const username = worksheet[XLSX.utils.encode_cell({ c: 0, r: i })].v;
  const password = worksheet[XLSX.utils.encode_cell({ c: 1, r: i })].v;
  const venue = worksheet[XLSX.utils.encode_cell({ c: 2, r: i })].v;
  const bookingTimeStartCell = worksheet[XLSX.utils.encode_cell({ c: 3, r: i })];
  const bookingTimeEndCell = worksheet[XLSX.utils.encode_cell({ c: 4, r: i })];
  delete bookingTimeStartCell.w;
  delete bookingTimeEndCell.w;
  bookingTimeStartCell.z = 'hh:mm:ss';
  bookingTimeEndCell.z = 'hh:mm:ss';
  XLSX.utils.format_cell(bookingTimeStartCell);
  XLSX.utils.format_cell(bookingTimeEndCell);
  const bookingTimeStart = bookingTimeStartCell.w;
  const bookingTimeEnd = bookingTimeEndCell.w;
  const usernameObj = order[username];
  const bookingTimes = (usernameObj && usernameObj.bookingTimes) || [];

  order[username] = {
    ...usernameObj,
    password,
    venue,
    bookingTimes: [...bookingTimes, [bookingTimeStart, bookingTimeEnd]]
  };
}

(async () => {
  const ordersArray = Object.entries(order);

  const bookingPromises = ordersArray.map(order => {
    const [username, user] = order;
    const { password, venue, bookingTimes } = user;

    return book(username, password, bookingDate, bookingTimes, venue);
  });

  console.log(await Promise.all(bookingPromises));
})();
