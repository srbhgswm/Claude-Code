% Logarithmic Spiral Plot
% Polar form: r = a * exp(b * theta)

a = 1;      % scale factor
b = 0.2;    % growth rate (controls how tightly the spiral winds)

theta = linspace(0, 6*pi, 1000);
r = a * exp(b * theta);

% Convert polar to Cartesian coordinates
x = r .* cos(theta);
y = r .* sin(theta);

figure;
plot(x, y, 'b-', 'LineWidth', 2);
hold on;
plot(0, 0, 'ro', 'MarkerSize', 8, 'MarkerFaceColor', 'r');  % origin
hold off;

axis equal;
grid on;
title(sprintf('Logarithmic Spiral  (a = %.1f, b = %.1f)', a, b));
xlabel('x');
ylabel('y');
