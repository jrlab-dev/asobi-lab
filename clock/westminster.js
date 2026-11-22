// ウェストミンスターの正時の旋律。各句の四音目のあとに一拍の間を置く。
const towerLeadSeconds = 25;
const towerStrikeInterval = 4.5;
const towerDecay = 4.8;

function towerStrikeCount(hour) {
    return hour % 12 || 12;
}

function ringWestminster(context, onEnded, {hour = null, firstStrikeIn = towerLeadSeconds} = {}) {
    const E = 329.63, Gs = 415.30, Fs = 369.99, B = 246.94;
    const phrases = [[E, Gs, Fs, B], [E, Fs, Gs, E], [Gs, E, Fs, B], [B, Fs, Gs, E]];
    const tower = hour !== null;
    const beat = tower ? 1.20 : 0.67;
    const now = context.currentTime;
    const firstStrike = now + firstStrikeIn;
    const start = tower ? firstStrike - towerLeadSeconds : now + 0.03;
    const output = context.createGain();
    output.connect(context.destination);
    // 低い唸り、打音の芯、短く消える金属音を別々に減衰させる。
    const partials = [
        { ratio: 0.5, level: 0.052, decay: 2.85 },
        { ratio: 1, level: 0.18, decay: 2.6 },
        { ratio: 1.2, level: 0.046, decay: 1.8 },
        { ratio: 2.005, level: 0.070, decay: 1.35 },
        { ratio: 2.98, level: 0.035, decay: 0.58 },
        { ratio: 4.17, level: 0.021, decay: 0.24 }
    ];
    const voices = new Set();
    let stopping = false;
    function strike(pitch, at, spectrum) {
        if (at < now - 0.005) return;
        at = Math.max(at, now);
        for (const { ratio, level, decay } of spectrum) {
            const voice = context.createOscillator();
            const envelope = context.createGain();
            voice.frequency.value = pitch * ratio;
            envelope.gain.setValueAtTime(0, at);
            envelope.gain.linearRampToValueAtTime(level, at + 0.012);
            envelope.gain.exponentialRampToValueAtTime(0.00005, at + decay);
            voice.connect(envelope);
            envelope.connect(output);
            voices.add(voice);
            voice.onended = () => {
                voice.disconnect();
                envelope.disconnect();
                voices.delete(voice);
                if (!voices.size) {
                    output.disconnect();
                    if (!stopping) onEnded();
                }
            };
            voice.start(at);
            voice.stop(at + decay + 0.02);
        }
    }
    phrases.forEach((phrase, bar) => {
        phrase.forEach((pitch, index) => strike(pitch, start + (bar * 5 + index) * beat, partials));
    });
    if (tower) {
        const greatBell = [
            {ratio: 0.5, level: 0.14, decay: towerDecay},
            {ratio: 1, level: 0.22, decay: 4.2},
            {ratio: 1.2, level: 0.075, decay: 3.2},
            {ratio: 2.004, level: 0.092, decay: 2.6},
            {ratio: 2.97, level: 0.048, decay: 1.45},
            {ratio: 4.18, level: 0.025, decay: 0.46}
        ];
        for (let index = 0; index < towerStrikeCount(hour); index++) {
            strike(E / 2, firstStrike + index * towerStrikeInterval, greatBell);
        }
    }
    return () => {
        if (stopping) return;
        stopping = true;
        const now = context.currentTime;
        output.gain.setTargetAtTime(0, now, 0.006);
        for (const voice of voices) voice.stop(now + 0.03);
    };
}
