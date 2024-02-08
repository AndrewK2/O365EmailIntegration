import Imap from "imap";
import Connection from "imap";

type Messages = any[];

export class ImapFetcher {
  constructor(private readonly emailAddress: string,
              private readonly accessToken: string,
              private readonly appConsole?: Console) {
  }

  public async Fetch(): Promise<Messages> {
    const console = this.appConsole;

    const build_XOAuth2_token = (user: string, access_token: string) => Buffer
      .from([`user=${user}`, `auth=Bearer ${access_token}`, '', ''].join('\x01'), 'utf-8')
      .toString('base64');

    const xoauth2 = build_XOAuth2_token(this.emailAddress, this.accessToken);

    console?.info("Using XOAuth2: " + xoauth2);

    const imap = new Imap({
      password: "",
      user: "",
      host: "outlook.office365.com",
      port: 993,
      tls: true,
      authTimeout: 15000,
      xoauth2: xoauth2,
      debug: (params: any) => {
        console?.log(params)
      }
    });

    this.appConsole?.info("Connecting to IMAP...")

    let emailArray: Messages = [];
    let flag = false;

    const email = async () => {
      imap.connect();

      return new Promise<any[]>((resolve, reject) => {
        imap.once('end', async function () {
          console?.log('Connection ended');
          resolve(emailArray);
          flag = true
        });
      })
    }

    function openInbox(cb: (error: Error, mailbox: Connection.Box) => void) {
      imap.openBox('INBOX', true, cb);
    }

    imap.once('ready', function () {
      try {
        emailArray = [];
        flag = false;

        openInbox(function (err, box) {
          if(err) {
            throw new Error("Error on mailbox open: " + err);
          }

          imap.search([['ALL'], ['SINCE', '1 Jan, 2024']],
            async function (err, uids) {
              if(err) {
                console?.log('Search error: ' + err);
                imap.end();
                return;
              }
              console?.debug("Search complete: ", uids);

              const fetch = imap.fetch(uids.sort().reverse().filter((u, i) => i < 10), {bodies: 'HEADER.FIELDS (FROM DATE)'});

              fetch.on('message', function (msg, seqno) {
                console?.log('Message #%d', seqno);

                msg.on('body', function (stream, info) {
                  let buffer = '';
                  stream.on('data', function (chunk) {
                    buffer += chunk.toString('utf8');
                  });

                  stream.once('end', function () {
                    let headers = Imap.parseHeader(buffer);
                    headers['UID'] = [seqno.toString()];

                    const emailFields: any = {};
                    for (const key in headers) {
                      emailFields[key] = headers[key][0];
                    }
                    emailArray.push(emailFields)
                  });
                });

                msg.once('end', function () {
                  console?.log(`(#${seqno}) Finished`);
                });
              });

              fetch.once('error', function (err) {
                console?.log('Fetch error: ' + err);
              });

              fetch.once('end', function () {
                console?.log('Done fetching all messages!');
                imap.end();
              });
            });
        });
      } catch (e) {
        console?.error(e);
      }
    });

    imap.once('error', function (err: any) {
      console?.log(err);
    });

    return await email();
  }
}
