/* 星の丘の文字板。閉位置と四分割時で同じ原画を使う。 */
const drawHillDial = (() => {
  const plate = document.createElement("canvas");
  plate.width = plate.height = 480;
  const g = plate.getContext("2d");
  g.translate(240, 240);
  g.scale(3, 3);
  function circle(ctx, x, y, r, color) {
    ctx.beginPath();
    ctx.arc(x, y, r, 0, 2 * Math.PI);
    ctx.fillStyle = color;
    ctx.fill();
  }
  function curve(ctx, points, color) {
    ctx.beginPath();
    ctx.moveTo(...points[0]);
    for (const p of points.slice(1)) ctx.bezierCurveTo(...p);
    ctx.closePath();
    ctx.fillStyle = color;
    ctx.fill();
  }
  function line(ctx, x, y, X, Y, width, color) {
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(X, Y);
    ctx.lineWidth = width;
    ctx.strokeStyle = color;
    ctx.lineCap = "round";
    ctx.stroke();
  }
  function star(ctx, x, y, r) {
    ctx.save();
    ctx.translate(x, y);
    ctx.beginPath();
    for (let i = 0; i < 8; i++) {
      let a = (i * Math.PI) / 4,
        len = i % 2 ? r * 0.28 : r;
      ctx.lineTo(Math.sin(a) * len, Math.cos(a) * len);
    }
    ctx.closePath();
    ctx.fillStyle = "#f2d48e";
    ctx.fill();
    ctx.restore();
  }
  let sky = g.createLinearGradient(0, -76, 0, 76);
  sky.addColorStop(0, "#15283f");
  sky.addColorStop(0.55, "#387678");
  sky.addColorStop(1, "#9db990");
  g.fillStyle = sky;
  g.fillRect(-80, -80, 160, 160);
  circle(g, -29, -34, 11, "#f7db94");
  circle(g, -25, -38, 10, "#233e51");
  for (const [x, y, r] of [
    [-43, -19, 2],
    [20, -49, 2],
    [42, -24, 2.7],
    [-4, -59, 2.5],
    [3, -29, 1.6],
    [53, 0, 2.2],
    [-62, -6, 2],
    [29, -12, 1.4],
  ])
    star(g, x, y, r);
  curve(
    g,
    [
      [-84, 16],
      [-52, -8, -20, 34, 12, 12],
      [46, -10, 60, -12, 84, -5],
      [84, 90, -82, 90, -84, 16],
    ],
    "#83ac96",
  );
  g.fillStyle = "#2c5254";
  g.fillRect(13, 1, 9, 11);
  curve(
    g,
    [
      [11, 2],
      [15, -8, 21, -8, 25, 2],
    ],
    "#f6d494",
  );
  line(g, 18, -5, 18, -9, 0.8, "#e8ca8a");
  curve(
    g,
    [
      [-84, 23],
      [-56, 4, -20, 7, 10, 28],
      [33, 45, 55, 16, 84, 16],
      [84, 90, -84, 91, -84, 23],
    ],
    "#386865",
  );
  curve(
    g,
    [
      [18, 11],
      [10, 14, 5, 18, 13, 22],
      [45, 35, -44, 46, -13, 78],
      [-6, 80, 5, 80, 11, 78],
      [-32, 47, 47, 37, 16, 22],
      [6, 17, 14, 14, 18, 11],
    ],
    "#dfbd7d",
  );
  for (const [x, y, h] of [
    [-53, 22, 19],
    [-40, 20, 13],
    [56, 29, 17],
    [64, 26, 10],
    [-62, 32, 12],
  ]) {
    line(g, x, y, x, y + h * 0.2, 0.9, "#244c4e");
    for (let k = 0; k < 3; k++) {
      g.beginPath();
      g.moveTo(x, y - h + k * h * 0.22);
      g.lineTo(x - h * (0.19 + 0.035 * k), y - h * 0.5 + k * h * 0.23);
      g.lineTo(x + h * (0.19 + 0.035 * k), y - h * 0.5 + k * h * 0.23);
      g.fillStyle = "#285555";
      g.fill();
    }
  }
  function ease(t) {
    t = Math.max(0, Math.min(1, t));
    return t * t * (3 - 2 * t);
  }
  function progress(t, a, b) {
    return ease((t - a) / (b - a));
  }
  function gear(c, x, y, r, angle) {
    c.save();
    c.translate(x, y);
    c.rotate(angle);
    c.beginPath();
    for (let i = 0; i < 48; i++) {
      const a = (i * Math.PI) / 24,
        rr = i % 4 < 2 ? r : r * 0.84;
      c.lineTo(rr * Math.cos(a), rr * Math.sin(a));
    }
    c.closePath();
    c.fillStyle = "#ac8958";
    c.fill();
    circle(c, 0, 0, r * 0.57, "#323e46");
    for (let i = 0; i < 3; i++) {
      line(
        c,
        0,
        0,
        r * 0.56 * Math.cos(i * 2.094),
        r * 0.56 * Math.sin(i * 2.094),
        1.7,
        "#ac8958",
      );
    }
    circle(c, 0, 0, 2, "#e5c888");
    c.restore();
  }
  function pet(c, x, t, right) {
    const enter =
      progress(t, right ? 2.55 : 1.8, right ? 3.55 : 2.75) -
      progress(t, 6.12, 6.75);
    const bow = progress(t, 5.4, 5.78) - progress(t, 5.9, 6.15);
    c.save();
    c.translate(x, 13 + 46 * (1 - enter));
    c.fillStyle = "#bb955a";
    c.fillRect(-16, 9, 32, 4);
    c.scale(1, 1 - 0.24 * bow);
    const fur = right ? "#9ccfe4" : "#efabc0";
    circle(c, 0, 3, 12, fur);
    c.save();
    c.translate(0, -6 + bow * 2);
    circle(c, -9, -9, 4.6, fur);
    circle(c, 9, -9, 4.6, fur);
    circle(c, 0, 0, 12, fur);
    for (const side of [-1, 1]) {
      circle(c, side * 4.4, -1.2, 1.45, "#413849");
      circle(c, side * 7, 2, 1.6, right ? "#7fb2cf" : "#dd8fa6");
    }
    c.beginPath();
    c.moveTo(-2, 2);
    c.quadraticCurveTo(0, 4.6, 2, 2);
    c.strokeStyle = "#53434c";
    c.lineWidth = 0.9;
    c.stroke();
    c.restore();
    for (const side of [-1, 1]) {
      const s = right ? 4.25 : 3.48;
      const wave = Math.sin(progress(t, s, s + 0.95) * 2 * Math.PI) * 0.78;
      c.save();
      c.translate(side * 9, 3);
      c.rotate(side * (0.25 + wave * (side === (right ? -1 : 1) ? 1 : 0.24)));
      line(c, 0, 0, side * 6, 2, 4.3, fur);
      c.restore();
    }
    c.restore();
  }

  const quadrants = [
    [-1, -1],
    [1, -1],
    [-1, 1],
    [1, 1],
  ];
  function drawPicture(c, photo) {
    if (!photo) {
      c.drawImage(plate, -80, -80, 160, 160);
      return;
    }
    c.fillStyle = "#000";
    c.fillRect(-80, -80, 160, 160);
    c.drawImage(photo.image, photo.x, photo.y, photo.width, photo.height);
  }
  return function drawHillDial(
    c,
    radius,
    shape,
    elapsed,
    background,
    reducedMotion,
  ) {
    c.save();
    const unit = radius / 100;
    c.scale(unit, unit);
    const widthScale = shape === "ellipse" ? 0.77 : 1;
    c.scale(widthScale, 1);
    let photo = null;
    if (background) {
      photo = {
        image: background.image,
        x: background.x / (unit * widthScale),
        y: background.y / unit,
        width: background.width / (unit * widthScale),
        height: background.height / unit,
      };
    }
    c.beginPath();
    c.arc(0, 0, 76, 0, 2 * Math.PI);
    c.clip();
    let t = elapsed;
    if (reducedMotion) t = t >= 1 && t < 7 ? 4 : 0;
    if (t <= 0 || t >= 8) {
      drawPicture(c, photo);
    } else {
      circle(c, 0, 0, 78, "#172e43");
      c.save();
      c.rotate(t * 0.27);
      for (let i = 0; i < 18; i++) {
        c.rotate(Math.PI / 9);
        star(c, 0, -56, 2 + (i % 3) * 0.6);
      }
      c.restore();
      gear(c, -7, -17, 13, t * 0.86);
      gear(c, 17, -17, 13, -t * 0.86 + Math.PI / 12);
      c.save();
      c.beginPath();
      c.rect(-66, -45, 132, 73);
      c.clip();
      pet(c, -30, t, false);
      pet(c, 30, t, true);
      c.restore();
      line(c, -66, 29, 66, 29, 2, "#ae935f");
      quadrants.forEach(([dx, dy], i) => {
        const travel =
          progress(t, 0.14 * i, 1.42 + 0.14 * i) -
          progress(t, 6.8 + 0.08 * i, 7.76 + 0.08 * i);
        c.save();
        c.translate(dx * 43 * travel, dy * 31 * travel);
        c.beginPath();
        c.rect(dx < 0 ? -80 : 0, dy < 0 ? -80 : 0, 80.06, 80.06);
        c.clip();
        drawPicture(c, photo);
        c.restore();
      });
    }
    c.restore();
  };
})();
