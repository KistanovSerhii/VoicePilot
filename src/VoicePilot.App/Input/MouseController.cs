using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using WindowsInput;
using WindowsInput.Events;

namespace VoicePilot.App.Input;

public class MouseController : IDisposable
{
    private readonly ILogger<MouseController> _logger;
    private CancellationTokenSource? _continuousMovementCts;
    private readonly object _movementLock = new();

    public MouseController(ILogger<MouseController> logger)
    {
        _logger = logger;
    }

    public void MoveBy(int deltaX, int deltaY)
    {
        Simulate.Events()
            .MoveBy(deltaX, deltaY)
            .Invoke()
            .GetAwaiter()
            .GetResult();
    }

    public void MoveTo(int x, int y)
    {
        Simulate.Events()
            .MoveTo(x, y)
            .Invoke()
            .GetAwaiter()
            .GetResult();
    }

    public void Click(string button, bool isDouble = false)
    {
        var buttonCode = button.ToLowerInvariant() switch
        {
            "left" => ButtonCode.Left,
            "right" => ButtonCode.Right,
            "middle" => ButtonCode.Middle,
            _ => ButtonCode.Left
        };

        var sequence = Simulate.Events();
        if (isDouble)
        {
            sequence.DoubleClick(buttonCode);
        }
        else
        {
            sequence.Click(buttonCode);
        }

        sequence.Invoke().GetAwaiter().GetResult();
    }

    public void StartContinuousMove(int deltaX, int deltaY, int intervalMilliseconds, CancellationToken cancellationToken)
    {
        lock (_movementLock)
        {
            StopContinuousMoveInternal();

            _continuousMovementCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            var linkedToken = _continuousMovementCts.Token;

            _ = Task.Run(async () =>
            {
                try
                {
                    while (!linkedToken.IsCancellationRequested)
                    {
                        Simulate.Events()
                            .MoveBy(deltaX, deltaY)
                            .Invoke()
                            .GetAwaiter()
                            .GetResult();

                        await Task.Delay(intervalMilliseconds, linkedToken).ConfigureAwait(false);
                    }
                }
                catch (OperationCanceledException)
                {
                    // expected when stopped
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Continuous mouse movement failed.");
                }
            }, linkedToken);
        }
    }

    public void StopContinuousMove()
    {
        lock (_movementLock)
        {
            StopContinuousMoveInternal();
        }
    }

    private void StopContinuousMoveInternal()
    {
        if (_continuousMovementCts is not null)
        {
            try
            {
                _continuousMovementCts.Cancel();
                _continuousMovementCts.Dispose();
            }
            catch (ObjectDisposedException)
            {
            }
            finally
            {
                _continuousMovementCts = null;
            }
        }
    }

    public void Dispose()
    {
        StopContinuousMove();
    }
}
