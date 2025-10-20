import smpp from 'smpp';

export class SmppClient {
    constructor(opts) {
        this.url = opts.url;
        this.system_id = opts.system_id;
        this.password = opts.password;
        this.senderId = opts.senderId;
        this.tps = Math.max(1, parseInt(opts.tps || 20, 10));
        this.connectTimeout = parseInt(opts.connectTimeout || 30000, 10);
        this.enquireMs = parseInt(opts.enquireMs || 15000, 10);
        this.reconnectMin = parseInt(opts.reconnectMin || 2000, 10);
        this.reconnectMax = parseInt(opts.reconnectMax || 30000, 10);

        this.session = null;
        this.connected = false;
        this.bound = false;
        this.queue = [];
        this.timer = null;
        this.backoff = this.reconnectMin;
        this.tickMs = Math.ceil(1000 / this.tps); // ~50ms for 20 TPS
    }

    connectAndBind() {
        return new Promise((resolve) => {
            const opts = {
                url: this.url,
                auto_enquire_link_period: this.enquireMs,
                debug: false,
                connectTimeout: this.connectTimeout
            };

            this.session = smpp.connect(opts, () => {
                this.connected = true;
                this._bind().then(resolve).catch(() => { });
            });

            this.session.on('error', (e) => {
                console.error('SMPP error:', e?.code || e?.message || e);
            });

            this.session.on('close', () => {
                console.warn('SMPP session closed.');
                this.connected = false;
                this.bound = false;
                this._stopSenderLoop();
                this._scheduleReconnect();
            });

            this.session.on('pdu', (pdu) => {
                if (pdu.command === 'deliver_sm') {
                    const isReceipt = (pdu.esm_class & 0x04) === 0x04;
                    const payload = pdu.short_message || pdu.message_payload;
                    console.log(isReceipt ? 'DLR:' : 'MO:', payload?.toString?.() || payload);
                    this.session.send(pdu.response());
                }
            });
        });
    }

    async _bind() {
        return new Promise((resolve, reject) => {
            if (!this.session) return reject(new Error('No session'));
            this.session.bind_transceiver({
                system_id: this.system_id,
                password: this.password
            }, (pdu) => {
                if (pdu.command_status === 0) {
                    console.log('✅ Bound as transceiver');
                    this.bound = true;
                    this.backoff = this.reconnectMin;
                    this._startSenderLoop();
                    resolve();
                } else {
                    console.error('❌ Bind failed with status', pdu.command_status);
                    this.session.close();
                    reject(new Error(String(pdu.command_status)));
                }
            });
        });
    }

    _scheduleReconnect() {
        const delay = this.backoff;
        console.log(`Reconnecting in ${delay} ms...`);
        setTimeout(() => this.connectAndBind().catch(() => { }), delay);
        this.backoff = Math.min(this.backoff * 2, this.reconnectMax);
    }

    _startSenderLoop() {
        if (this.timer) return;
        this.timer = setInterval(() => this._drainOne(), this.tickMs);
    }
    _stopSenderLoop() {
        if (this.timer) clearInterval(this.timer);
        this.timer = null;
    }

    _drainOne() {
        if (!this.bound || !this.session) return;
        const job = this.queue.shift();
        if (!job) return;

        const { to, text, resolve, reject, options } = job;

        const submit = {
            destination_addr: to,
            dest_addr_ton: 1,
            dest_addr_npi: 1,
            source_addr: this.senderId,
            source_addr_ton: 5,
            source_addr_npi: 0,
            registered_delivery: 1,
            data_coding: 0,
            short_message: text,
            ...(options || {})
        };

        this.session.submit_sm(
            submit,
            (pdu) => {
                if (pdu.command_status === 0) {
                    resolve({ message_id: pdu.message_id, to, text });
                } else {
                    reject(new Error(`submit_sm status ${pdu.command_status}`));
                }
            },
            () => { },
            (failure) => reject(new Error(`Socket write failed: ${failure?.message || failure}`))
        );
    }

    send(to, text, options = null) {
        return new Promise((resolve, reject) => {
            this.queue.push({ to, text, options, resolve, reject });
        });
    }
}
